
# Installed Applications Parser

A Node.js script to parse installed applications on macOS and export the data to an Excel file.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/phuongwd/installed-applications-parser.git
   cd installed-applications-parser
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

## Generating the Installed Applications File

To create the `Installed_Applications.txt` file, run the following command in your Terminal:

```bash
system_profiler SPApplicationsDataType > Installed_Applications.txt
```

This command gathers information about all installed applications and saves it to `Installed_Applications.txt`.

## Usage

Once you have the `Installed_Applications.txt` file in your project directory, run the following command to parse the data and export it to an Excel file:

```bash
npm start
```

The application will output the data to `Installed_Applications.xlsx`.

## License

This project is licensed under the MIT License. See [LICENSE.md](LICENSE.md).
