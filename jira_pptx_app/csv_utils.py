import csv
from io import StringIO

def generate_csv_from_issues(issues, fields, fixversion_ce, fixversion_ee):
    csv_data = StringIO()
    writer = csv.writer(csv_data)
    
    # Write headers
    writer.writerow(fields + ['CE', 'EE'])
    
    for issue in issues:
        summary = getattr(issue.fields, 'summary', None)
        if not summary:
            summary = "No Summary Provided"

        data = []
        for field in fields:
            value = None
            if field == 'Issue Key':
                value = issue.key
            elif field.lower() == 'summary':  # Ensure case-insensitivity
                value = summary
            elif field.lower() == 'fixversion':
                value = issue.fields.fixVersions[0].name if issue.fields.fixVersions else ''
            # Handle other standard fields in a similar manner if needed...
            else:
                # Handle custom fields
                value = getattr(issue.fields, field, None)

            data.append(value if value is not None else '')

        ce_tick = '✓' if fixversion_ce in [fv.name for fv in issue.fields.fixVersions] else '-'
        ee_tick = '✓' if fixversion_ee in [fv.name for fv in issue.fields.fixVersions] else '-'
        writer.writerow(data + [ce_tick, ee_tick])

    # Now, write the contents of csv_data to 'output.csv'
    with open('output.csv', 'w', newline='') as csvfile:
        csvfile.write(csv_data.getvalue())
