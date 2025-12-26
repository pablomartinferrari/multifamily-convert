# 1. SharePoint Site Setup and Document Libraries

This section covers setting up your SharePoint Online environment with the necessary site structure, document libraries, and permissions for the Excel processing solution.

## Prerequisites

- SharePoint Online tenant with admin access
- Site Collection Administrator permissions
- Modern SharePoint site (Communication or Team site)

## Step 1: Create or Configure SharePoint Site

### Option A: Create New Communication Site

1. Navigate to your SharePoint admin center: `https://[tenant]-admin.sharepoint.com`
2. Go to **Sites** → **Active sites**
3. Click **Create** → **Communication site**
4. Configure:
   - **Site name**: "Excel Processing Center"
   - **Site address**: `excel-processing`
   - **Template**: Choose "Topic" or "Blank"
   - **Description**: "Upload Excel files and generate automated reports"

### Option B: Use Existing Modern Site

If you prefer to use an existing site, ensure it has modern pages enabled and you have design permissions.

## Step 2: Create Document Libraries

You need four document libraries:
- **Input Files**: Where users upload Excel files
- **Report 1**: For the first generated report
- **Report 2**: For the second generated report
- **Report 3**: For the third generated report

### Creating Document Libraries

1. Navigate to your SharePoint site
2. Click **New** → **Document library**
3. Create the following libraries:

#### Input Files Library
```
Name: Input Files
Description: Upload your Excel files here for processing
```

#### Report Libraries
Create three report libraries:

```
Name: Monthly Summary Reports
Description: Generated monthly summary reports

Name: Data Quality Reports
Description: Data validation and quality analysis reports

Name: Trend Analysis Reports
Description: Historical trend analysis and forecasting reports
```

## Step 3: Configure Library Settings

### Input Files Library Configuration

1. Go to **Input Files** library → **Settings** (gear icon) → **Library settings**
2. Under **General Settings**:
   - **Versioning settings**: Enable versioning (optional but recommended)
   - **Require check out**: No (to allow concurrent uploads)

3. Under **Permissions and Management**:
   - **Manage files which have no checked-in version**: Allow (recommended)

### Report Libraries Configuration

For each report library:

1. **Library settings** → **Versioning settings**:
   - Create major versions: Yes
   - Keep drafts: No

2. **Advanced settings**:
   - Content approval: No
   - Document versioning: Create major versions

## Step 4: Set Up Folder Structure (Optional)

Create logical folder organization in your libraries:

### Input Files Library Structure
```
/Input Files
├── /Batch_2024_Q1/
├── /Batch_2024_Q2/
└── /Pending_Processing/
```

### Report Libraries Structure
Each report library should have:
```
/Monthly Summary Reports
├── /2024/
│   ├── /Q1/
│   ├── /Q2/
│   └── /Q3/
└── /2025/
```

## Step 5: Configure Permissions

### Recommended Permission Structure

1. **Site Owners Group**: Full control
2. **Excel Processors Group**: Contribute permissions on Input Files, Read on Reports
3. **Report Viewers Group**: Read permissions on all libraries

### Setting Up Permission Groups

1. Go to **Site settings** → **People and groups**
2. Create new groups:

#### Excel Processors Group
- **Name**: Excel Processors
- **Description**: Users who can upload and process Excel files
- **Permissions**: Contribute on Input Files library, Read on Report libraries

#### Report Viewers Group
- **Name**: Report Viewers
- **Description**: Users who can view generated reports
- **Permissions**: Read on all libraries

### Applying Permissions

1. For **Input Files** library:
   - Break inheritance from site
   - Grant Contribute permissions to Excel Processors group
   - Remove default members group

2. For **Report libraries**:
   - Break inheritance
   - Grant Read permissions to Report Viewers group
   - Grant Contribute permissions to Excel Processors group

## Step 6: Create Modern Page with Document Library Web Parts

1. Navigate to your site home page
2. Click **Edit** (or create new page with **+ New** → **Page**)
3. Add web parts:

### Add Document Library Web Parts

1. Click **+** to add web part
2. Search for "Document library"
3. Add four Document library web parts:
   - One for Input Files
   - One for each Report library

### Configure Web Parts

For each Document library web part:
1. Click **Edit web part** (pencil icon)
2. Select the appropriate library
3. Configure view:
   - **View**: All Documents
   - **Layout**: List or Compact List
   - Show command bar: Yes

### Layout Recommendations

Use a 2-column layout:
- **Left Column**: Input Files library
- **Right Column**: Stacked report libraries

## Step 7: Enable Custom Scripting (if needed)

If you'll be deploying SPFx web parts to this site:

1. Go to **Site settings** → **Site collection features**
2. Activate: **SharePoint Server Publishing Infrastructure** (if not already active)

## Step 8: Test Basic Functionality

1. Upload a test Excel file to Input Files library
2. Verify you can access all report libraries
3. Test different user permissions by logging in as different users

## Folder Structure Summary

Your SharePoint site should now have:

```
Excel Processing Center/
├── Input Files/           # Document library for uploads
├── Monthly Summary Reports/    # Generated reports
├── Data Quality Reports/       # Generated reports
├── Trend Analysis Reports/     # Generated reports
├── Site Pages/
│   └── Excel Processing.aspx   # Main processing page
└── Site Assets/          # For SPFx web parts (created later)
```

## Next Steps

With your SharePoint site configured, proceed to [creating the SPFx web part](./02-spfx-webpart.md) that will provide the processing button and file selection interface.

## Troubleshooting

### Common Issues

1. **Cannot create document libraries**: Ensure you have sufficient permissions
2. **Web parts not available**: Ensure you're using a modern site
3. **Permission inheritance issues**: Use "Check Permissions" in site settings to diagnose

### Permission Testing

Create test accounts and verify:
- Excel Processors can upload to Input Files
- Excel Processors can view reports
- Report Viewers can only read (not upload)

