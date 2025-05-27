import re
from datetime import datetime

def replace_placeholders(text, email_content):
    # Replace company name and address in the agreement paragraph
    text = re.sub(r'dated _____________, is between _______________________________________________ \("Customer"\), with offices at ________________________________________________ _____________, USA',
                  'dated _____________, is between XYZ Co. ("Customer"), with offices at 330, Flatbush Avenue, Brooklyn, New York 11238, USA',
                  text)

    # Replace scope section
    scope_pattern = r'\[Add the services that are within the scope of this SOW\]\n\[If this SOW is for an extension, add the following sentences. If not, delete it.\nThe initial SOW for this task was executed in <Month> <Year>. This SOW is for an extension of the <date/scope/number of hours>.\]'
    text = re.sub(scope_pattern, email_content['scope'], text)
    text = re.sub(r'xx', email_content['hours'], text)

    # Replace assumptions section
    assumptions_pattern = r'•\s*\[Add any assumptions/limitations here. If there aren’t any, delete the entire box.\]'
    if email_content['assumptions']:
        text = re.sub(assumptions_pattern, f'• {email_content["assumptions"]}', text)
    else:
        text = re.sub(r'Assumptions and Limitations:\n•\s*\[Add any assumptions/limitations here. If there aren’t any, delete the entire box.\]\n?', '', text)

    # Replace resource assignment section
    resources_pattern = r'\[If there are specific resources/roles working for this SOW, mention them along with the following sentence:\nABC will assign the following resources to complete this project.\]'
    text = re.sub(resources_pattern, email_content['resources'], text)

    # Replace [year] with actual year
    current_year = str(datetime.now().year)
    text = re.sub(r'\[year\]', current_year, text)

    # Preserve "Name :" exactly
    text = re.sub(r'Name\s*:', 'Name :', text)

    # Keep "THE MASTER AGREEMENT AND" in original color
    text = re.sub(r'(THE MASTER AGREEMENT AND)', r'<span style="color:red;">\1</span>', text)

    return text

# Example usage
if __name__ == "__main__":
    email_content = {
        'scope': 'This SOW includes development, testing, and deployment of the new analytics dashboard. The initial SOW for this task was executed in January 2025. This SOW is for an extension of 40 hours.',
        'hours': '40',
        'assumptions': 'Customer will provide access to required systems and documentation.',
        'resources': 'ABC will assign John Doe (Developer) and Jane Smith (QA Analyst) to complete this project.'
    }

    with open('sow_document.txt', 'r') as file:
        sow_text = file.read()

    updated_text = replace_placeholders(sow_text, email_content)

    with open('updated_sow_document.txt', 'w') as file:
        file.write(updated_text)

    print("SOW document updated successfully.")
