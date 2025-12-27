# Data Processing Requirements

This document outlines the requirements and business logic for processing lead-based paint inspection data within this workspace.

## 1. Overview
The solution automates the ingestion, normalization, and quality analysis of XRF (X-ray fluorescence) lead-based paint inspection data stored in Excel format. It transforms raw machine outputs into structured SharePoint list items suitable for regulatory reporting.

## 2. Data Ingestion
Files are uploaded to the `XRF Inspection Files` document library and processed via a SharePoint Framework (SPFx) command set.

### 2.1 File Format
- **Format**: Excel (.xlsx, .xls)
- **Source**: XRF machine exports.
- **Mandatory Metadata**: Before processing, each file must have a `Job Number` (Title), `Property Name`, and `File Type` (Units or Common Areas).

### 2.2 Column Mapping
The processor must handle varied column headers (case-insensitive):
- **Reading**: `Reading`, `Shot #`, `Shot`
- **Component**: `Component`
- **Lead Measurement**: `Pb`, `Lead`, `Pb mg/cm2`
- **Result**: `Result` (Pos/Neg)
- **Location**: `Side`, `Color`, `Substrate`, `Condition`, `Room Number`, `Room Type`, `Floor`, `Unit`, `Building`

## 3. Data Transformation Rules

### 3.1 Lead Status Determination
The system calculates or validates the `Lead Status` based on the measurement (`Pb` value) and a threshold (default: **1.0 mg/cm²**):
- **Positive**: `Pb` >= 1.0 OR Excel Result = "Pos"
- **Negative**: `Pb` < 1.0 AND Excel Result = "Neg"
- **Excluded**: `Pb` is null/empty OR `Component` = "CALIBRATE"

### 3.2 Unit/Area Classification
Each shot must be classified as either a "Unit" or "Common Area":
- **Common Areas**: Set automatically if the `File Type` is "Common Areas" or if `Room Type` matches known common area patterns (e.g., Hallway, Lobby).
- **Units**: Set if `File Type` is "Units". Use the `Unit` column if available; otherwise, use `Room Number` as the unit identifier.
- **Unknown**: Flag for review if classification cannot be determined.

### 3.3 Component Normalization
Raw component names (e.g., "door jamb", "Door Jamb", "Door-Jamb") must be normalized using a lookup against the `Component Master` list:
1. Exact match against `Title`.
2. Match against `Common Variations` field.
3. If no match, flag as a "Component Naming" issue.

### 3.4 Calibration Shots
- Shots where the component is "CALIBRATE" are flagged as `Is Calibration = True`.
- These shots are excluded from lead percentage calculations and statistical reports.

## 4. AI-Powered Quality Assurance
Azure OpenAI (GPT-4) is integrated via Power Automate to perform advanced data quality checks:
- **Consistency**: Identify naming variations that escaped basic normalization.
- **Anomalies**: Detect patterns like measurements clustering exactly at the 1.0 threshold.
- **Accuracy**: Validate if a shot classified as a "Unit" actually belongs in a "Common Area" based on `Room Type`.
- **Explanations**: Generate natural language explanations for flagged issues to assist inspectors.

## 5. Regulatory & Reporting Requirements (Michigan/HUD)

### 5.1 The 40-Shot Minimum
Per Michigan/HUD requirements for lead-based paint reporting:
- **Requirement**: A component must have **at least 40 individual shots** to be eligible for component-level averaging.
- **Implementation**:
    - Reports must exclude averages for components with < 40 shots.
    - These components should be flagged with a warning: "Insufficient shots for statistical averaging."

### 5.2 Lead Percentage Calculation
The primary metric for reporting is the Lead Positive Percentage:
$$ \text{Lead \%} = \left( \frac{\text{Count of Positive Shots}}{\text{Total Valid Shots}} \right) \times 100 $$
*Note: Valid shots exclude calibration shots and those with quality errors.*

## 6. Data Integrity
- **Referential Integrity**: Every `Inspection Shot` must link back to its `Source File` and `Component Master` entry.
- **Audit Trail**: The original raw component name (`Component Raw`) must be preserved alongside the normalized version.
- **Change Tracking**: File status must transition through `Uploaded` -> `Processing` -> `Complete` (or `Error`).

