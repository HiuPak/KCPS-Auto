# KCPS_Auto

A Windows automation tool that simplifies repetitive configuration tasks by programmatically controlling an external application using UI Automation.

---

## 📌 Overview

KCPS_Auto is a C# automation tool designed to reduce repetitive manual configuration tasks in the application *KanCollePlayerSimulator*. Due to limitations in the free version of the simulator, settings must be reconfigured on every startup. This project automates that process, reducing several minutes of repetitive clicking to a single action.

The project was built for personal use and as a learning exercise in Windows UI Automation and structured automation scripting.

---

## ✨ Features

- Automates UI interactions based on predefined scripts
- Reduces repetitive configuration tasks to a single execution
- Supports script selection for flexible automation behaviour
- Detects and interacts with running external processes

---

## 🧩 Architecture

The solution consists of two applications:

1. **KCPS_Auto**  
   The main automation program responsible for controlling the target application using a consistent configuration.

2. **KCPS_ScriptSelector**  
   An optional helper tool that allows users to select different automation scripts before execution.

This separation allows configuration logic to evolve independently from execution logic.

---

## 🛠 Technologies Used

- C#
- .NET Framework 4.7.2
- System.Windows.Automation
- JSON (System.Text.Json)
- File I/O
- Windows desktop application development

---

## 🧪 Testing & Development Approach

The project was developed with a focus on reliability and repeatability. Automation logic was tested by validating UI element discovery, process detection, and script execution under different application states. Configuration data is stored and deserialized from JSON to ensure scripts remain readable and easy to modify.

While this is a personal project, I followed a test-driven mindset by validating expected automation behaviour incrementally and ensuring changes did not break existing functionality. Logging and controlled execution steps were used to verify correctness before extending features.

---

## ▶ How to Run

### Prerequisites
- Windows OS
- .NET Framework 4.7.2
- Visual Studio 2019 or later
- *KanCollePlayerSimulator* running locally

### Steps
1. Clone the repository
2. Open the `.sln` file in Visual Studio
3. Build the solution
4. Ensure `KanCollePlayerSimulator.exe` is running
5. If required, update the process name in `App.config` (e.g. version changes)
6. Run the application (`Ctrl + F5`)

---

## 📈 What This Project Demonstrates

- Practical use of Windows UI Automation APIs
- Automation of real-world repetitive workflows
- Structured handling of external processes
- JSON-based configuration and scripting
- Clear separation of concerns and maintainable design
- Careful handling of automation reliability and user interaction

---

## ⚠ Disclaimer

This project is for educational and personal use only. It is not affiliated with or endorsed by the developers of KanCollePlayerSimulator.
