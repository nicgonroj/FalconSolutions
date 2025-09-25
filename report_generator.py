"""Generate purchase order report without relying on SQL Server temporary tables.

This module reproduces the behaviour of the original SQL script that populated
`#TEMPPOS` by iterating over analysis filters and executing the
`dbo.task_PODetail` stored procedure.  The implementation has been moved to
Python so the intermediate state is held in memory, avoiding the need to write
into `tempdb`.

Typical usage from the command line::

    python report_generator.py \
        --connection "Driver={ODBC Driver 18 for SQL Server};Server=..." \
        --company-id 4849 \
        --group-code BSC \
        --output report.xlsx

The script connects to the configured database, fetches the analysis filters
matching the provided group code, executes the stored procedure for every
filter, aggregates the results and finally writes them to an Excel workbook.
"""
from __future__ import annotations

import argparse
import logging
import os
from dataclasses import dataclass
from typing import Iterable, List, Optional

import pandas as pd
import pyodbc


LOGGER = logging.getLogger(__name__)


@dataclass
class Config:
    """Runtime configuration for the report generation process."""

    connection_string: str
    company_id: int
    group_code: Optional[str]
    output_path: str


ANALYSIS_FILTER_QUERY = """
SELECT AnalysisFilterId, AnalysisFilterName
FROM AnalysisFilter WITH (NOLOCK)
WHERE CompanyId = ?
  AND AnalysisFilterTypeId = 3
  AND AnalysisFilterName LIKE ?
ORDER BY AnalysisFilterName;
"""

PODetail_SP = """
EXEC dbo.task_PODetail @CompanyId = ?, @AnalysisFilterId = ?, @StatusCodeId = 6;
"""

FINAL_COLUMNS: List[str] = [
    "AnalysisFilterName",
    "PurchaseOrderCode",
    "PurchaseOrderName",
    "SentDate",
    "TenderCode",
    "STenderClientFullName",
    "STenderBuyingUnitFullName",
    "CurrencyType",
    "TotalPO",
    "CompetitorName",
    "PurchaseOrderLink",
    "PurchaseOrderDocumentLink",
    "ONU",
    "ClientProductDesc",
    "CompetitorProductDesc",
    "Qty",
    "NetAmount",
    "Discount",
    "Charges",
    "Taxes",
    "Total",
    "CurrencyTypeItem",
    "ItemNbr",
    "SRegionName",
    "RUT",
    "MPProductId",
    "ProductTypeDescription",
    "ProductBrandName",
    "ProductModel",
    "ProductMeasure",
    "ONUCategoryCode",
    "ONUCategoryDescription",
    "BuyingContactFullName",
    "BuyingContactEmail",
    "BuyingContactPhone",
]


def build_group_pattern(group_code: Optional[str]) -> str:
    """Return the LIKE pattern used to filter analysis filters.

    The T-SQL script prefixes the provided group name with ``|`` before applying
    the ``LIKE`` clause.  We mirror the behaviour so the Python implementation
    returns the same results as the SQL version.
    """

    if not group_code:
        return "%"
    return f"%|{group_code}%"


def fetch_analysis_filters(connection: pyodbc.Connection, config: Config) -> pd.DataFrame:
    """Retrieve the analysis filters that should be processed."""

    pattern = build_group_pattern(config.group_code)
    LOGGER.debug("Fetching analysis filters with pattern %s", pattern)
    return pd.read_sql_query(
        ANALYSIS_FILTER_QUERY,
        connection,
        params=[config.company_id, pattern],
    )


def fetch_purchase_order_details(
    connection: pyodbc.Connection, company_id: int, analysis_filter_id: int
) -> pd.DataFrame:
    """Execute the stored procedure and return the details as a DataFrame."""

    LOGGER.debug(
        "Executing task_PODetail for company_id=%s, analysis_filter_id=%s",
        company_id,
        analysis_filter_id,
    )
    return pd.read_sql_query(PODetail_SP, connection, params=[company_id, analysis_filter_id])


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure that the dataframe contains the expected set of columns.

    The stored procedure already returns most of the fields required by the
    final report.  However, to keep the pipeline robust we add any missing
    columns so the export phase can rely on a consistent schema.
    """

    missing_columns = [column for column in FINAL_COLUMNS if column not in df.columns]
    for column in missing_columns:
        LOGGER.debug("Adding missing column %s", column)
        df[column] = pd.NA

    # Ensure CurrencyTypeItem matches the behaviour of the original SQL script.
    if "CurrencyTypeItem" in df.columns and df["CurrencyTypeItem"].isna().all():
        if "CurrencyType" in df.columns:
            df["CurrencyTypeItem"] = df["CurrencyType"]
    return df[FINAL_COLUMNS]


def generate_report(connection: pyodbc.Connection, config: Config) -> pd.DataFrame:
    """Create the consolidated report as a DataFrame."""

    filters_df = fetch_analysis_filters(connection, config)
    if filters_df.empty:
        LOGGER.warning("No analysis filters found for the specified parameters")
        return pd.DataFrame(columns=FINAL_COLUMNS)

    LOGGER.info("Processing %d analysis filters", len(filters_df))
    results: List[pd.DataFrame] = []

    for _, row in filters_df.iterrows():
        filter_id = int(row["AnalysisFilterId"])
        filter_name = row["AnalysisFilterName"]
        details_df = fetch_purchase_order_details(connection, config.company_id, filter_id)
        if details_df.empty:
            LOGGER.info("No purchase orders found for analysis filter %s", filter_name)
            continue

        details_df["AnalysisFilterName"] = filter_name
        details_df["CurrencyTypeItem"] = details_df.get("CurrencyTypeItem", details_df.get("CurrencyType"))
        results.append(normalise_columns(details_df))

    if not results:
        LOGGER.warning("No purchase order information was retrieved")
        return pd.DataFrame(columns=FINAL_COLUMNS)

    combined = pd.concat(results, ignore_index=True)
    LOGGER.info("Generated report with %d rows", len(combined))
    return combined


def write_to_excel(df: pd.DataFrame, output_path: str) -> None:
    """Persist the report to an Excel workbook."""

    LOGGER.debug("Writing report to %s", output_path)
    df.to_excel(output_path, index=False)


def parse_args(argv: Optional[Iterable[str]] = None) -> Config:
    parser = argparse.ArgumentParser(description="Generate the purchase order report")
    parser.add_argument(
        "--connection",
        dest="connection_string",
        default=os.getenv("FALCON_SQL_CONNECTION"),
        help="ODBC connection string for the SQL Server database",
    )
    parser.add_argument("--company-id", type=int, required=True, help="Company identifier")
    parser.add_argument(
        "--group-code",
        dest="group_code",
        default=os.getenv("FALCON_GROUP_CODE"),
        help="Optional group code used to filter analysis filters",
    )
    parser.add_argument(
        "--output",
        dest="output_path",
        default="purchase_order_report.xlsx",
        help="Destination Excel file path",
    )

    args = parser.parse_args(argv)

    if not args.connection_string:
        parser.error("A connection string must be supplied via --connection or FALCON_SQL_CONNECTION")

    return Config(
        connection_string=args.connection_string,
        company_id=args.company_id,
        group_code=args.group_code,
        output_path=args.output_path,
    )


def main(argv: Optional[Iterable[str]] = None) -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    config = parse_args(argv)

    LOGGER.info("Starting report generation")
    with pyodbc.connect(config.connection_string) as connection:
        df = generate_report(connection, config)
        write_to_excel(df, config.output_path)

    LOGGER.info("Report saved to %s", config.output_path)


if __name__ == "__main__":
    main()
