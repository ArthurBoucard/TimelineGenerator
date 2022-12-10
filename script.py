import openpyxl
import requests

# GitHub API base URL
API_BASE_URL = 'https://api.github.com'

# GitHub project board owner and repository
PROJECT_BOARD_OWNER = '<OWNER>'
PROJECT_BOARD_REPO = '<REPO>'

# GitHub API access token (required for private repositories)
# Generate a personal access token at https://github.com/settings/tokens
ACCESS_TOKEN = '<ACCESS_TOKEN>'

# GitHub project board ID
PROJECT_BOARD_ID = '<PROJECT_BOARD_ID>'

# GitHub milestone number
MILESTONE_NUMBER = '<MILESTONE_NUMBER>'

# Excel workbook and worksheet to create
EXCEL_WORKBOOK = 'timeline.xlsx'
EXCEL_WORKSHEET = 'Timeline'


def main():
    # get the list of issues in the specified milestone
    issues = get_milestone_issues(PROJECT_BOARD_OWNER, PROJECT_BOARD_REPO, MILESTONE_NUMBER)

    # create an Excel workbook and worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = EXCEL_WORKSHEET

    # write the issue names and durations to the worksheet
    for issue in issues:
        worksheet.append([issue['title'], issue['duration']])

    # save the workbook
    workbook.save(EXCEL_WORKBOOK)


def get_milestone_issues(owner, repo, milestone_number):
    """
    Get the list of issues in the specified milestone.
    """
    # construct the URL for the GitHub API request
    url = f'{API_BASE_URL}/repos/{owner}/{repo}/issues?milestone={milestone_number}&state=all'

    # send the request and parse the JSON response
    response = requests.get(url, headers={'Authorization': f'token {ACCESS_TOKEN}'})
    issues = response.json()

    # extract the issue names and durations from the response
    results = []
    for issue in issues:
        results.append({
            'title': issue['title'],
            'duration': issue['duration']
        })

    return results


if __name__ == '__main__':
    main()
