import os
import json
import datetime
from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession
import openpyxl

# Scopes for Google Analytics API
SCOPES = ["https://www.googleapis.com/auth/analytics.readonly"]

# Load service account JSON directly from the environment variable.
service_account_info = json.loads(os.environ.get("GA_SERVICE_ACCOUNT_JSON"))
credentials = service_account.Credentials.from_service_account_info(
    service_account_info, scopes=SCOPES
)

# Retrieve GA property ID from environment variable.
PROPERTY_ID = os.environ.get("GA_PROPERTY_ID")
if not PROPERTY_ID:
    raise ValueError("GA_PROPERTY_ID environment variable not set.")

def compute_last_week_date_range():
    """
    Compute last week’s Monday and Sunday.
    In JavaScript, Sunday=0, but in Python, Monday=0 and Sunday=6.
    If today is Sunday (weekday() == 6), then:
      last_sunday = today - 7 days and last_monday = today - 13 days.
    Otherwise:
      last_sunday = today - (weekday() + 1) days and last_monday = last_sunday - 6 days.
    """
    today = datetime.date.today()
    if today.weekday() == 6:  # Sunday in Python
        last_sunday = today - datetime.timedelta(days=7)
        last_monday = today - datetime.timedelta(days=13)
    else:
        last_sunday = today - datetime.timedelta(days=(today.weekday() + 1))
        last_monday = last_sunday - datetime.timedelta(days=6)
    return last_monday.strftime("%Y-%m-%d"), last_sunday.strftime("%Y-%m-%d")

def get_ga4_report(payload, credentials):
    """Call the GA4 API to get the report."""
    url = f"https://analyticsdata.googleapis.com/v1beta/properties/{PROPERTY_ID}:runReport"
    session = AuthorizedSession(credentials)
    response = session.post(url, json=payload)
    response.raise_for_status()
    return response.json()

def save_to_excel(data, filename="analytics_data.xlsx"):
    """
    Write the GA4 API response data to an Excel file.
    The GA4 runReport response includes 'dimensionHeaders', 'metricHeaders', and 'rows'.
    """
    wb = openpyxl.Workbook()
    ws = wb.active

    # Prepare headers from response metadata.
    dimension_headers = [header["name"] for header in data.get("dimensionHeaders", [])]
    metric_headers = [header["name"] for header in data.get("metricHeaders", [])]
    headers = dimension_headers + metric_headers
    ws.append(headers)

    # Write each row: GA4 response rows include 'dimensionValues' and 'metricValues'.
    for row in data.get("rows", []):
        dimensions = [dim.get("value", "") for dim in row.get("dimensionValues", [])]
        metrics = [met.get("value", "") for met in row.get("metricValues", [])]
        ws.append(dimensions + metrics)

    wb.save(filename)
    print(f"Data successfully saved to {filename}")

def main():
    # Compute date range for last week (Monday to Sunday)
    start_date, end_date = compute_last_week_date_range()
    print(f"Date range: {start_date} to {end_date}")

    # Construct payload with dimensions, metrics, and filters.
    payload = {
        "dateRanges": [
            {"startDate": start_date, "endDate": end_date}
        ],
        "dimensions": [
            {"name": "customEvent:product"},
            {"name": "date"}
        ],
        "metrics": [
            {"name": "eventCount"}
        ],
        "dimensionFilter": {
            "andGroup": {
                "expressions": [
                    {
                        "filter": {
                            "fieldName": "pageLocation",
                            "stringFilter": {
                                "matchType": "CONTAINS",
                                "value": "?utm_source=companion&utm_medium=pva&utm_campaign=commercial_chatbot"
                            }
                        }
                    },
                    {
                        "filter": {
                            "fieldName": "pageLocation",
                            "stringFilter": {
                                "matchType": "CONTAINS",
                                "value": "en-us"
                            }
                        }
                    },
                    {
                        "filter": {
                            "fieldName": "eventName",
                            "stringFilter": {
                                "matchType": "EXACT",
                                "value": "complete_sub"
                            }
                        }
                    },
                    {
                        "notExpression": {
                            "orGroup": {
                                "expressions": [
                                    {
                                        "filter": {
                                            "fieldName": "customEvent:issue_category",
                                            "stringFilter": {
                                                "matchType": "EXACT",
                                                "value": "脅威に関する問題"
                                            }
                                        }
                                    },
                                    {
                                        "filter": {
                                            "fieldName": "customEvent:issue_category",
                                            "stringFilter": {
                                                "matchType": "EXACT",
                                                "value": "Threat Issue"
                                            }
                                        }
                                    }
                                ]
                            }
                        }
                    }
                ]
            }
        }
    }

    # Retrieve report.
    report = get_ga4_report(payload, credentials)
    print("Report retrieved successfully.")

    # Save report data to an Excel file.
    save_to_excel(report, filename="analytics_data.xlsx")

if __name__ == "__main__":
    main()
