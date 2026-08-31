# 🩸 Blood Bank & Donor Management System (BBMS)

![Frontend](https://img.shields.io/badge/Frontend-Visual_Basic_6.0-blue)
![Backend](https://img.shields.io/badge/Backend-Oracle_SQL-red)
![Project](https://img.shields.io/badge/Project-BCA_Final_Year-success)

---

## 📌 Project Overview
The **Blood Bank Management System (BBMS)** is a desktop-based application developed to digitalize and streamline the complete workflow of a blood bank. It efficiently manages donor registration, medical validation, lab testing, component processing, inventory tracking, and blood issuance.

This project was developed as a **Final Year BCA Project**, focusing on **real-world problem solving, database design, and efficient data handling**.

---

## 🎯 Objectives
- Manage donor and patient records efficiently  
- Maintain real-time blood stock  
- Ensure safe and validated blood issuance  
- Generate accurate and dynamic reports  

---

## 🛠️ Technology Stack
- **Frontend:** Visual Basic 6.0 (VB6)  
- **Backend:** Oracle Database (SQL)  
- **Database Connectivity:** ADODB (ActiveX Data Objects)  
- **Reporting:** DataEnvironment & MS DataReports  

---

## ✨ Core Modules & Features

### 🧑‍🤝‍🧑 Donor Management & Medical Validation
- Donor registration with complete demographic details  
- Medical validation (Weight, BP, Pulse, Temperature, HB Level)  
- Automatic classification: **FIT / UNFIT / DEFERRED**  
- Donor ID card generation  

---

### 🧪 Lab Testing & Component Processing
- Unique Sample ID generation  
- Blood testing (HIV, HCV, Malaria)  
- Result status: **PASS / FAIL**  
- Component separation:
  - RBC  
  - Plasma  
  - Platelets  
- Expiry tracking for each component  

---

### 📦 Blood Stock Management
- Real-time inventory tracking  
- Status management:
  - AVAILABLE  
  - ISSUED  
  - EXPIRED  
- Auto-update on blood issue  

---

### 🏥 Patient Requests & Blood Issuance
- Patient request and hospital management  
- Cross-matching validation (**MATCH / NOT MATCH**)  
- Automated billing system  
- Payment tracking  

---

### 📊 Reporting System
- Dynamic report generation  
- Search by ID and date range  
- Multi-table SQL JOIN queries  

---

## 🤖 RedBot – Smart Chat Assistant
The system includes an inbuilt chatbot named **RedBot**, designed to improve usability and user interaction.

### 🧠 Features:
- Guides users through different modules  
- Helps in navigating system functionalities  
- Provides quick answers to basic queries  
- Reduces dependency on manual help  

### 🎯 Purpose:
RedBot acts as a **virtual assistant**, making the system easier to use, especially for beginners.

### 🚀 Future Enhancements:
- AI-based chatbot (NLP integration)  
- Voice interaction support  
- Smart recommendations  

---

## 🔐 Advanced Concepts Used
- **Database Normalization**  
- **Complex SQL JOINs**  
- **Soft Delete Mechanism** (STATUS = Active / Inactive)  
- **Case-Insensitive Queries** using `UPPER()`  
- **Duplicate Prevention using SQL logic**  

---

## 🚀 Setup Instructions

### 📌 Prerequisites
- Visual Basic 6.0 IDE  
- Oracle Database (10g/11g or above)  
- Oracle ODBC/OLEDB Drivers  

---

### ⚙️ Installation Steps

1. **Database Setup**
   - Open Oracle SQL Developer / SQL*Plus  
   - Execute the database script file  

2. **Project Setup**
   - Open `Project1.vbp` in VB6  
   - Update database connection in `CONN` module  

3. **Run Application**
   - Press `F5` to run the project  

---

## 📚 Learning Outcomes
- Real-world database design  
- Writing optimized SQL queries  
- Desktop application development  
- Data validation and reporting systems  

---

## 🔮 Future Scope
- Convert into Web-Based Application  
- Add Login & Authentication System  
- Improve UI/UX Design  
- Integrate Notification System (SMS/Email)  
- Upgrade RedBot to AI-powered chatbot  

---

## 👨‍💻 Author
**Rachit Kumar Singh**  
🎓 BCA Final Year Student  

---

## 📬 Feedback
Feedback and suggestions are always welcome!  
Feel free to connect and share your thoughts.

---

## ⭐ Support
If you like this project, give it a ⭐ on GitHub!