const fs = require("fs");
const xlsx = require("xlsx");

// Define the input and output file paths
const inputFilePath = "Installed_Applications.txt";
const outputExcelPath = "Installed_Applications.xlsx";

// Function to determine usage category based on the last modified date
function determineUsageCategory(lastModified) {
  const lastModifiedDate = new Date(lastModified.split(",")[0]); // Extract date part
  const currentDate = new Date();
  const daysDifference = Math.floor(
    (currentDate - lastModifiedDate) / (1000 * 60 * 60 * 24)
  ); // Convert to days

  if (daysDifference <= 60) return "Frequently";
  if (daysDifference <= 120) return "Occasionally";
  return "Rarely";
}

// Function to determine the license type based on the "Obtained from" string
function determineLicenseType(obtainedFrom) {
  const lowerCaseString = obtainedFrom.toLowerCase();

  if (
    lowerCaseString.includes("free") ||
    lowerCaseString.includes("freeware")
  ) {
    return "Freeware";
  } else if (
    lowerCaseString.includes("paid") ||
    lowerCaseString.includes("purchase")
  ) {
    return "Paid";
  } else if (
    lowerCaseString.includes("open source") ||
    lowerCaseString.includes("source")
  ) {
    return "Open Source";
  }
  return "";
}

function isInstalledInApplications(location) {
  return /^(\/Applications\/|\/Users\/[^\/]+\/Applications\/)/.test(location);
}

function extractValue(text, prefix) {
  const regex = new RegExp(`${prefix}:\\s*(.*)`);
  const match = text.match(regex);
  return match ? match[1].trim() : null;
}

function extractDeveloperID(text) {
  // Regular expression to capture the name and ID in parentheses
  const regex =
    /(Developer ID Application|Apple Development):\s*([^()]+)\s*\(([A-Z0-9]+)\)/;

  // Execute the regular expression on the text
  const match = text.match(regex);

  // Return the captured developer name and ID if found
  return match ? `${match[2].trim()} (${match[3]})` : null;
}


function extractLastModified(data) {
  const regex = /Last Modified:\s*(\d{1,2}\/\d{1,2}\/\d{2}),\s*(\d{2}:\d{2})/;
  const match = data.match(regex);
  return match ? `${match[1]}, ${match[2]}` : null;
}

// Function to check if the input file exists and is valid
function checkInputFile(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error(
      `Warning: The file "${filePath}" does not exist. Please generate it first.`
    );
    process.exit(1); // Exit the program with a failure code
  }

  const fileStats = fs.statSync(filePath);
  if (fileStats.size === 0) {
    console.error(
      `Warning: The file "${filePath}" is empty. Please ensure it contains valid data.`
    );
    process.exit(1); // Exit the program with a failure code
  }
}

async function parseInstalledApplications(filePath) {
  const data = fs.readFileSync(filePath, "utf-8");
  const applications = [];
  const lines = data
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line);
  let currentApp = null;

  lines.forEach((line) => {
    if (line.endsWith(":") && !line.includes("Version")) {
      // Finalize the previous application entry if it exists
      if (currentApp) {
        if (isInstalledInApplications(currentApp.Location)) {
          applications.push(currentApp);
        }
      }
      currentApp = { ID: applications.length + 1, Name: line.slice(0, -1) };
    } else if (line.startsWith("Version:")) {
      currentApp.Version = extractValue(line, "Version") || "N/A";
    } else if (line.startsWith("Signed by:")) {
      const signedBy = extractValue(line, "Signed by");
      currentApp.Vendor = extractDeveloperID(signedBy) || signedBy;
    } else if (line.startsWith("Last Modified:")) {
      const lastModified = extractLastModified(line);
      currentApp.Used = determineUsageCategory(lastModified);
    } else if (line.startsWith("Obtained from:")) {
      currentApp["License Type"] = determineLicenseType(
        extractValue(line, "Obtained from")
      );
    } else if (line.startsWith("Location:")) {
      currentApp.Location = extractValue(line, "Location");
    }
  });

  // Finalize the last application entry
  if (currentApp && isInstalledInApplications(currentApp.Location)) {
    applications.push(currentApp);
  }

  return applications;
}

async function exportToExcel(data, filePath) {
  const workbook = xlsx.utils.book_new();

  // Map data to include custom headers
  const customData = data.map((app) => ({
    ID: app.ID,
    Name: app.Name,
    Version: app.Version,
    Vendor: app.Vendor,
    Used: app.Used,
    "License Type": app["License Type"],
    OS: 'macOS'
  }));

  const worksheet = xlsx.utils.json_to_sheet(customData);
  xlsx.utils.book_append_sheet(workbook, worksheet, "Applications");
  xlsx.writeFile(workbook, filePath);
  console.log(`Data exported successfully to ${filePath}`);
}

async function main() {
  try {
    // Check if the input file exists and is valid
    checkInputFile(inputFilePath);

    const applicationsData = await parseInstalledApplications(inputFilePath);

    // Sort applications by name in alphabetical order
    applicationsData.sort((a, b) => a.Name.localeCompare(b.Name));

    // Update IDs after sorting
    applicationsData.forEach((app, index) => {
      app.ID = index + 1; // Assign new ID based on the sorted position
    });

    await exportToExcel(applicationsData, outputExcelPath);
  } catch (error) {
    console.error("Error:", error.message);
  }
}

main();
