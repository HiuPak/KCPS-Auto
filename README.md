# KCPS_Auto

This is a program that automatically control the setting of other program "KanCollePlayerSimulator(https://github.com/KanaHayama/KanCollePlayerSimulator Free to use, 16 hours every start, setting for the program will not saved for free version)".

---

## 📌 Description

This application recude my daily 5 minutes clicking work in the specified program into one button
The application is for self usage purpose only
It also helps me to get use the basic usage of System.Windows.Automation

There are two programe
1. KCPS_Auto - main programe, have to use the same setting everytime
2. KCPS_ScriptSelector - sub program, optional to chooose script

---

## ✨ Features
Automate clicking base on written script

---

## 🛠 Technologies Used
- C#
- .NET Framework 4.7.2
- File I/O
- System.Windows.Automation
- System.Text.Json
- JSON serialization & deserialization

---

## ▶ How to Run
### Prerequisites
- Windows OS
- .NET Framework 4.7.2
- Visual Studio 2019 or later

### Steps
1. Clone the repository
2. Open the `.sln` file in Visual Studio
3. Build the solution
4. Make sure the app "KanCollePlayerSimulator.exe" already started
5. Change process name of "KanCollePlayerSimulator.exe" in App.config if necessary (mainly version number)
6. Run the application (Ctrl + F5)