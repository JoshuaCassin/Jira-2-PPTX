from jira import JIRA

def fetch_jira_data(jira_server_url, email, api_key, jql_query):
    try:
        # Initialize Jira client with provided credentials
        jira = JIRA(jira_server_url, basic_auth=(email, api_key))

        # Initialize variables for pagination
        max_results = 50  # Adjust as needed, but be aware of Jira server limitations
        start_at = 0
        all_issues = []

        while True:
            # Retrieve a batch of Jira data using the provided query
            issues = jira.search_issues(jql_query, startAt=start_at, maxResults=max_results)
            
            # Append retrieved issues to the list
            all_issues.extend(issues)
            
            # Check if we have fetched all available issues
            if len(issues) < max_results:
                break
            
            # Update the start_at for the next iteration
            start_at += len(issues)

        return all_issues
    except Exception as e:
        raise RuntimeError(f"Error fetching Jira data: {str(e)}")
