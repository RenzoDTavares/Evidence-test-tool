# 🚀 Test Assistant: Automated Evidence & Azure DevOps Integration

## 📖 Overview

The **Test Assistant** is a high-performance Quality Engineering tool designed to eliminate documentation bottlenecks in large-scale enterprise environments.

Conceptualized and beta-tested during a consultancy engagement at **EY** for **Bradesco** (one of Latin America’s largest banks), this application automates the bridge between local test execution and cloud-based test management.

By leveraging the **Azure DevOps REST API**, the tool transforms manual reporting into a streamlined, automated workflow—ensuring audit compliance, traceability, and operational governance.

---

## 📉 The Problem: Documentation Fragmentation

In large financial projects, test evidence generation is often inconsistent and inefficient:

- **No Standard Format:** QA Analysts create evidence using Excel, Word, or raw image folders.
- **Manual Overhead:** Copying metadata (User Stories, IDs, Steps) from Azure DevOps is time-consuming and error-prone.
- **Traceability Gaps:** Mapping screenshots to specific test steps becomes a major administrative burden.

---

## ✨ Key Features

- **Single Source of Truth**  
  Automatically retrieves metadata from Azure DevOps, including User Stories, Test Cases, Environments, and linked Bugs.

- **Intelligent Image-to-Step Mapping**  
  Correlates screenshots with test steps based on execution order.

- **Standardized Templating**  
  Uses `.docx` templates with dynamic placeholders to ensure consistent and professional reports.

- **Automated Upload & Attachment**  
  Generated evidence is automatically attached back to the corresponding Test Case in Azure DevOps.

- **Secure Credential Management**  
  Supports Personal Access Token (PAT) handling with secure storage and validation.

---

## 🛠️ Technical Architecture

The application follows a decoupled, service-oriented architecture:

### Core Components

1. **AzureDevOpsService**  
   Handles REST API communication, test step extraction (XML), and attachment uploads.

2. **DocumentService**  
   Responsible for `.docx` generation, placeholder replacement, and image processing.

3. **QAController**  
   Orchestrates business logic and multi-threaded execution for performance and UI responsiveness.

---

### 🧰 Tech Stack

- Python  
- Azure DevOps REST API  
- XML Parsing  
- python-docx  
- Multi-threading  

---

## 🚀 How to Use

### 1. Prerequisites

- Python 3.10+
- A valid **Azure DevOps Personal Access Token (PAT)** with:
  - `Work Items: Read & Write`

---

### 2. Configuration

Update the `config.json` file:

{
  "ORGANIZATION_URL": "https://dev.azure.com/YourOrgName",
  "API_TIMEOUT": 10
}

---

### 3. Execution Steps

1. Launch the Application  
2. Token Setup  
3. Select Directory  
4. Enter Metadata  
5. Generate Report  

---

## 📈 Real-World Impact (Business Case)

During its deployment at EY/Bradesco, the tool delivered:

- Efficiency: Reduced reporting time from ~15 minutes to under 2 minutes per test case
- Consistency: Achieved 100% template compliance
- Approval: Validated by leadership, QAs and stakeholders

---

## 🎥 Watch it in Action

To better understand how the **Test Assistant** streamlines the QA workflow, I have prepared a comprehensive video demonstration. [cite_start]In this video, I walk through a real-use case, showing the transition from manual screenshot capturing to fully automated Azure DevOps documentation.

> **[Link to YouTube Video]**(https://youtu.be/_eaD4kBm-ck)

### What's covered in the demo:
* **The Problem Statement:** Why inconsistent evidence in large projects like Bradesco leads to technical debt.
* **Setup & Security:** How to generate and register a Personal Access Token (PAT) with the correct scopes.
* **Live Generation:** Fetching metadata (User Stories, Test Steps, and Bugs) directly from the "Evidence Project" in Azure.
* **Automated Attachment:** Verifying the generated `.docx` file linked directly to the Test Case attachments.

---

## 🤝 Contributing


Contributions are welcome, especially for:
- Jira/Xray integration
- PDF export features

