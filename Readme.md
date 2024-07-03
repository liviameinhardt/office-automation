# Office Automation

This project aims to automate various office tasks using Python. It provides a set of utilities and functions to automate emails, powerpoints and word documents.

## Installation

To use this project, follow these steps:

1. Clone the repository: `git clone https://github.com/your-username/office-automation.git`
2. Install the required dependencies: `pip install -r requirements.txt`

## Features

- Generate PowerPoint presentations with customizable [templates](https://www.wikihow.com/Edit-a-PowerPoint-Template)
- Generate Word documents
- Generate and send automated Outlook emails

## Usage

To get started, import the `office_templates` module and use the provided functions to generate office documents. 
A few examples:

### PowerPoint Presentation
```python
from office_templates import  presentation

#Create a presentation from an image FOLDER:
presentation("C:\MyPath\my_folder",title="MyPresentation")

# Create a presentation from a image LIST:
presentation(image=["img1.png","img2.png"],path="C:\MyPath\my_folder",title="MyPresentation")
```

### Outlook Email
```python
from office_templates import  email

email(body="Hi all, please check the images attached.",subject='Monthly reports',attachments='.../Reports',to_email='client@gmail.com; myboss@gmail.com')
```

## Contributing

This is a very simple and limited code. I only implemented what I personally needed, but improvements are always welcome!
If you have any ideas or suggestions, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more information.
