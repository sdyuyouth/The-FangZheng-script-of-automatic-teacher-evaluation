# The FangZheng Script of Automatic Teacher Evaluation

This is an automated script for evaluating teachers on the FangZheng system, designed to streamline the process of providing feedback on teaching performance.Teaching evaluation is often at the end of the semester, in order to save students in the evaluation of teaching troubles, this script came into being.The script will evaluate each teacher with the highest score combination and includes a check function to avoid accidentally giving bad reviews to teachers.

## Features

- **Automated Evaluation:** Automatically log in to the educational affairs system, and automatically evaluate, submit, and exit until all teachers have been evaluated.
- **Selenium WebDriver:** Utilizes Selenium for browser automation to interact with web elements.Progress can be seen in real time, and most anti-reptile measures can be circumvented.
- **Excel Integration:** Reads account credentials from an Excel file for multi-user support.

## Prerequisites

- Python 3.x
- Selenium WebDriver（Edge webdriver）
- An Excel file with account credentials

## Installation

1. Clone the repository:
git clone https://github.com/your-username/The-FangZheng-script-of-automatic-teacher-evaluation.git

2. Install the required Python packages:
pip install selenium pandas urllib3 time random

3. Download the appropriate EdgeDriver for your browser and place it in the project directory or set the path in the script.

## Usage

1. Prepare an Excel file named `pingjiao.xlsx` with the following columns:
- `username`: The account username.
- `password`: The account password.

2. Run the script:
python main.py


## Contributing

Contributions are welcome! Please follow these steps to contribute:
1. Fork the repository.
2. Create a new branch for your changes.
3. Make your changes and test them thoroughly.
4. Submit a pull request with a detailed description of your changes.

## License

This project is open source and available under the [MIT License](LICENSE).

## Disclaimer

This script is for educational purposes only. Use it responsibly and ensure you are complying with your institution's policies on automated interactions with web systems.
