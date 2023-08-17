import logging
import os

from jira import JIRA
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Sharepoint conf
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_URL = "https://epam.sharepoint.com/sites/TestingCC-Dev"
LIST_TITLE = "TCC Satisfaction Survey"

# Jira credentials
JIRA_URL = "https://jira.epam.com/jira"
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_PASS = os.getenv("JIRA_PASS")

# Constants
# JQL_QUERY = 'project="EPAM Testing Competency Center" and resolution in ("Done") and resolved  >= -7d' # TODO: Uncomment
JQL_QUERY = 'project="EPAM Testing Competency Center" and reporter in ("Olga_Marchenko@epam.com", "Ivan_Bozhko@epam.com") and resolved  >= -7d'
CUSTOMER_PROJECT_FIELD_NAME = "Customer/project"
NOT_AVAILABLE = "N/A"

logging.basicConfig(level=logging.INFO)


def connect_to_jira(url, login, path):
    jira = JIRA(
        basic_auth=(login, path),
        options={
            'server': url
        }
    )
    return jira


def authenticate_sharepoint():
    ctx = None
    try:
        ctx_auth = AuthenticationContext(SITE_URL)
        if ctx_auth.acquire_token_for_app(client_id=CLIENT_ID, client_secret=CLIENT_SECRET):
            ctx = ClientContext(SITE_URL, ctx_auth)
        else:
            logging.error("Failed to authenticate with SharePoint.")
    except Exception as e:
        logging.error(f"Failed to authenticate with SharePoint: {str(e)}")
    return ctx


def create_list_item(context, list_title, item_properties):
    try:
        target_list = context.web.lists.get_by_title(list_title)
        context.load(target_list)
        context.execute_query()

        target_list.add_item(item_properties)
        context.execute_query()
        logging.info("New item created successfully.")
    except Exception as e:
        logging.error(f"Failed to create new item: {e}")


def fetch_issues(jira, jql_query):
    jira_items = jira.search_issues(jql_query)
    return jira_items


def get_detailed_issue(jira, issue_key):
    detailed_issue = jira.issue(issue_key)
    return detailed_issue


def get_customer_project_field_id(jira, field_name):
    fields = jira.fields()
    customer_project_field = next((field for field in fields if field['name'] == field_name), None)
    if customer_project_field is None:
        print(f"Could not find custom field with name '{field_name}'")
        exit()
    return customer_project_field['id']


def get_customer_project_value(issue, field_id):
    try:
        customer_project_value = getattr(issue.fields, field_id)
    except AttributeError as e:
        logging.warning(f"Could not find custom field '{field_id}' in the issue {issue.key}. Error: {e}")
        customer_project_value = None
    return customer_project_value


def main():
    jira = connect_to_jira(JIRA_URL, JIRA_EMAIL, JIRA_PASS)

    jira_items = fetch_issues(jira, JQL_QUERY)
    logging.info(f"Jira issues was found: {jira_items}")

    for issue in jira_items:
        detailed_issue = get_detailed_issue(jira, issue.key)

        context = authenticate_sharepoint()
        if not context:
            logging.error("Failed to authenticate with SharePoint.")
            return

        components = detailed_issue.fields.components
        component = components[0].name if components else NOT_AVAILABLE

        customer_project_field_id = get_customer_project_field_id(jira, CUSTOMER_PROJECT_FIELD_NAME)
        customer_project_value = get_customer_project_value(issue, customer_project_field_id)

        # Create a new list item
        item_properties = {
            'Requestor': detailed_issue.fields.reporter.displayName if detailed_issue.fields.reporter else NOT_AVAILABLE,
            'Summary': detailed_issue.fields.summary if detailed_issue.fields.summary else NOT_AVAILABLE,
            'Component': component,
            'Customer': customer_project_value if customer_project_value else NOT_AVAILABLE,
            'JiraID': detailed_issue.key if detailed_issue.key else NOT_AVAILABLE,
            'Assignee': detailed_issue.fields.assignee.displayName if detailed_issue.fields.assignee else NOT_AVAILABLE,
        }
        logging.info(f"Creating list item for issue {detailed_issue.key} with properties: {item_properties}")
        create_list_item(context, LIST_TITLE, item_properties)
        logging.info(f"List item created successfully for issue {detailed_issue.key}")


if __name__ == '__main__':
    main()
