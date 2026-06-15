# Earnings-call-risk-tool
AI Risk Scanner for Earnings Calls
Why this project?
Analysts and investors spend a lot of time reading long earnings‑call transcripts and reports to pull out the key risks, which is repetitive, easy to rush and hard to do consistently under time pressure. I built this tool to see how AI and simple automation could make that process faster and more structured, while still keeping analysts in control of the judgement.

By combining Excel (where finance teams already work) with a lightweight Python + OpenAI backend, the project shows how large language models can be used in a practical, auditable way: turning unstructured text into a clear risk log with categories, evidence and follow‑up questions that fit naturally into equity and credit research, audit planning and risk discussions.

Overview
This project is an AI‑assisted risk analysis tool that combines an Excel workbook with a Python backend to help equity and credit analysts turn long earnings‑call excerpts into a structured risk register.

Given a section of an earnings call or annual report, the tool identifies the main risks, classifies them (macro/policy, credit, capital, operational, regulatory/compliance) and writes them back into Excel as clean, analyst‑ready rows.

Features
Excel‑first workflow

Input sheet where the user pastes an excerpt from an earnings call or report.

Risk Log sheet that stores one row per risk with fixed columns:

Risk Title, Category, Description, Why It Matters, Evidence, Follow‑up Question, Severity, Source.

Python + OpenAI backend

Reads the input text from Excel, sends it to an LLM with a prompt tailored for risk extraction, and receives structured JSON back.

Parses the JSON and writes each risk into the Risk Log sheet, appending below any existing rows.

Automatically resizes the Excel table to include new rows so filters and formatting stay intact.

Risk taxonomy and constraints

Categories restricted to: Macro/Policy, Credit, Capital, Operational, Regulatory/Compliance.

Severity must be one of: Low, Medium, High.

Evidence is always a short quote or paraphrased line from the source text.

Output limited to the 5–7 most important risks for that excerpt.

Files in this repository
risk_scanner.xlsx – Excel front‑end with:

Input sheet (paste transcript text into cell B4).

Risk Log sheet with the risk table and sample output based on HSBC’s Q1 2025 earnings call.

Instructions sheet explaining the basic workflow.

risk_scanner_table_safe.py – Python backend that:

Loads risk_scanner.xlsx.

Reads the input text from Input!B4.

Calls the OpenAI API with a structured prompt and fixed risk schema.

Parses the JSON response and appends risks into the Risk Log sheet.

Updates the Excel table range to cover all populated rows.

Project-Overview.docx – Project notes describing the motivation, MVP design and first worked example.

How it works
User preparation (in Excel)

Paste an earnings‑call excerpt or report section into the Input sheet (cell B4).

Python pipeline

The script risk_scanner_table_safe.py loads the workbook and reads the input text.

It builds a prompt that instructs the model to:

Extract 5–7 key risks.

Use the fixed risk taxonomy and severity levels.

Return only valid JSON with exactly the required keys.

The script calls the OpenAI Responses API (gpt-4o-mini by default) and retrieves the model output.

It cleans the response (stripping code fences, optional “json” labels) and parses the JSON into Python objects.

Each risk is written as a new row into Risk Log, and the associated Excel table range is resized to include the new rows.

Analyst review (back in Excel)

The user reviews the generated risk log, filters/sorts as needed, and uses the follow‑up questions to guide further analysis or management Q&A.

Setup
Clone this repository

bash
git clone https://github.com/<your-username>/<your-repo-name>.git
cd <your-repo-name>
Create and activate a virtual environment (optional but recommended)

bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
Install dependencies

bash
pip install openpyxl openai
Set your OpenAI API key

On macOS/Linux:

bash
export OPENAI_API_KEY="your_api_key_here"
On Windows (PowerShell):

powershell
$env:OPENAI_API_KEY="your_api_key_here"
Usage
Open risk_scanner.xlsx.

Go to the Input sheet and paste an earnings‑call or report excerpt into cell B4.

Save and close the workbook (or at least make sure it’s not open in a way that locks the file).

Run the Python script from the project folder:

bash
python risk_scanner_table_safe.py
When the script prints Workbook updated successfully (table-safe)., reopen risk_scanner.xlsx and check the Risk Log sheet.

New risks should appear as additional rows in the log, inside the formatted table.

Customisation
Model and source label

In risk_scanner_table_safe.py, you can change:

MODEL (default: "gpt-4o-mini") to any compatible OpenAI model.

SOURCE_LABEL to match the document you’re analysing (e.g. "Company X FY2025 annual report").

Table selection

By default, the script uses the first table it finds on the Risk Log sheet.

To target a specific table, set TABLE_NAME in the script to the table’s name from Excel.

Risk schema

The headers and categories are defined in HEADERS and in the prompt.

If you adjust them, update both the Python script and the Excel template so they stay in sync.

Limitations & future ideas
The tool currently processes one text excerpt at a time and expects it in Input!B4.

It relies on the OpenAI API, so usage requires an active API key and internet connection.

Future improvements could include:

A simple Excel button or macro that triggers the Python script for non‑technical users.

Batch processing of multiple calls and a comparison view across different companies or quarters.

More granular risk categories tailored to specific sectors (e.g. banks vs corporates).

