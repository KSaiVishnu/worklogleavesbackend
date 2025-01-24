const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

const cors = require("cors");
app.use(cors());
app.use(express.json());

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const filePath = req.file.path;
    const workbook = XLSX.readFile(filePath);

    const timelog = XLSX.utils.sheet_to_json(workbook.Sheets["Timelog"]);
    const leaves = XLSX.utils.sheet_to_json(workbook.Sheets["Leaves"]);
    function convertTimeToMinutes(timeStr) {
      const [hours, minutes] = timeStr.split(":");
      return parseInt(hours) * 60 + parseInt(minutes);
    }

    function convertDateToISO(dateStr) {
      const [day, month, year] = dateStr.split("-");
      return `${year}-${day}-${month}`;
    }

    function calculateTotalTime(timelogData) {
      const userDayLogs = [];
      timelogData.forEach((entry) => {
        const { User, Date } = entry;
        const dailyLog = entry["Daily Log"];
        if (dailyLog && Date) {
          const formattedDate = convertDateToISO(Date);
          const dailyLogInMinutes = convertTimeToMinutes(dailyLog);

          let logEntry = userDayLogs.find(
            (log) => log.User === User && log.Date === formattedDate
          );

          if (!logEntry) {
            logEntry = { User, Date: formattedDate, TotalMinutes: 0 };
            userDayLogs.push(logEntry);
          }

          logEntry.TotalMinutes += dailyLogInMinutes;
        }
      });
      return userDayLogs;
    }

    const leaveLookup = new Map();
    leaves.forEach((leave) => {
      const excelStartDate = new Date(1899, 11, 31);
      let x = new Date(excelStartDate.getTime() + leave.LeaveStartDate * 86400000);
      const startDate = new Date(x);
      const isHalfDay = leave.NumberOfDays === 0.5;

      for (let i = 0; i < Math.ceil(leave.NumberOfDays); i++) {
        const date = new Date(startDate);
        date.setDate(startDate.getDate() + i);

        const key = `${leave.EmployeeName}-${date.toISOString().split("T")[0]}`;
        leaveLookup.set(key, isHalfDay ? "half" : "full");
      }
    });

    const userDayLogs = calculateTotalTime(timelog);
    // console.log(leaveLookup);

    const filteredUsers = [];
    userDayLogs.forEach((entry) => {
      const dateKey = `${entry.User}-${entry.Date}`;
      const dailyLogInMinutes = entry.TotalMinutes;
      const leaveType = leaveLookup.get(dateKey);
      // console.log(entry,dateKey,leaveType);

      if (!leaveType && dailyLogInMinutes < 360) {
        filteredUsers.push(entry);        
      } else if (
        leaveType === "half" &&
        dailyLogInMinutes < 180
      ) {
        filteredUsers.push(entry);
      }
      
    });

    const outputSheet = XLSX.utils.json_to_sheet(filteredUsers);
    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "FilteredUsers");

    const outputFilePath = path.join(__dirname, "FilteredUsers.xlsx");
    XLSX.writeFile(outputWorkbook, outputFilePath);

    fs.unlinkSync(filePath);

    res.download(outputFilePath, "FilteredUsers.xlsx", () => {
      fs.unlinkSync(outputFilePath);
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing file.");
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () =>
  console.log(`Server running on http://localhost:${PORT}`)
);
