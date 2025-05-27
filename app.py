import re
from datetime import datetime

def replace_placeholders(text, email_content):
    # Replace company name and address in the agreement paragraph
    text = re.sub(
        r'dated _____________, is between _______________________________________________ \("Customer"\), with offices at ________________________________________________ _____________, USA',
        'dated _____________, is between XYZ Co. ("Customer"), with offices at 330, Flatbush Avenue, Brooklyn, New York 11238, USA',
        text
    )

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

    # Keep "THE MASTER AGREEMENT AND" in red
    text = re.sub(r'(THE MASTER AGREEMENT AND)', r'<span style="color:red;">\1</span>', text)

    return text
