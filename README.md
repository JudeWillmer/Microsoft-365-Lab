# Lab 1 – Microsoft Outlook Overview

This lab demonstrates common Microsoft Outlook administrative tasks performed by an IT Support Technician. The scenarios focus on mailbox management, user configuration, Microsoft Exchange features, calendar administration, and Outlook troubleshooting.

---

## Scenario 1 – Block Sender / Junk Email

**Purpose**

Learn how to block unwanted senders and manage Junk Email settings to help users control spam and recover legitimate email if accidentally blocked.

**Location**

Home → More (...) → Block

**Result**

Successfully accessed the Block Sender and Junk Email options from Outlook and reviewed the available spam management settings.

**Screenshot**
<img width="1918" height="971" alt="Image 1" src="https://github.com/user-attachments/assets/a2157501-e9ad-4017-a2d8-93a6f83e3d0a" />

## Scenario 2 – Global Address List (Address Book)

**Purpose**

Access the Global Address List (GAL) to locate users, verify contact information, and understand how Outlook integrates with Microsoft Exchange.

**Location**

Home → Address Book

**Result**

Successfully opened the Global Address List, searched for users, and viewed Exchange contact information.

**Screenshot**
<img width="1918" height="970" alt="Image 2" src="https://github.com/user-attachments/assets/e4118b5b-b63c-408f-9f87-4e06e797055f" />

## Scenario 3 – Download Offline Address Book

**Purpose**

Update the Offline Address Book to synchronize the latest Global Address List from Microsoft Exchange.

**Location**

Send/Receive → Send/Receive Groups → Download Address Book

**Result**

Successfully accessed the Offline Address Book download settings and verified synchronization options.

**Screenshot**
<img width="1918" height="971" alt="Image 3" src="https://github.com/user-attachments/assets/f6aab81f-efec-46e8-b640-5d5c3e058390" />

## Scenario 4 – Outlook View Configuration

**Purpose**

Configure common Outlook view settings used during Help Desk troubleshooting, including disabling Focused Inbox and adjusting the Reading Pane.

**Location**

View → Show Focused Inbox / Layout → Reading Pane

**Result**

Successfully reviewed Outlook view settings, disabled Focused Inbox, and verified Reading Pane configuration options.

**Screenshot**
<img width="1918" height="971" alt="Image 4" src="https://github.com/user-attachments/assets/99fa414a-faed-4cbf-ac05-ce1d22ef72fb" />

## Scenario 5 – Calendar Permissions

**Purpose**

Review calendar sharing permissions and understand how Outlook controls user access to calendars within Microsoft Exchange.

**Location**

Calendar → Share → Calendar Properties → Permissions

**Result**

Successfully opened Calendar Properties and reviewed the available permission levels for calendar sharing.

**Screenshot**
<img width="1918" height="970" alt="Image 5" src="https://github.com/user-attachments/assets/d10d304a-5ffc-4baa-b389-117af3119a76" />

## Scenario 6 – Delegate Access

**Purpose**

Configure Delegate Access to allow another user to manage calendar items and receive meeting requests on behalf of the mailbox owner.

**Location**

File → Account Settings → Delegate Access

**Result**

Successfully added a delegate and reviewed available permission levels for Calendar, Tasks, Inbox, Contacts, and Notes.

**Screenshot**
<img width="1918" height="971" alt="Image 6" src="https://github.com/user-attachments/assets/12707e16-434e-44aa-a05f-63205f0175ed" />

# Lab 2 – Office 365 Troubleshooting

## Objective

Troubleshoot common Microsoft 365 application issues using a structured troubleshooting approach. Learn how to identify common causes of Office application problems, isolate faults using techniques such as Safe Mode and COM Add-ins, check for stuck background processes, repair corrupted Office files, and use Microsoft 365 repair options before considering a full reinstallation.

---

# 2A – Office Applications Troubleshooting

## Scenario 1 – Office COM Add-ins

**Purpose**

Review installed COM Add-ins to identify third-party extensions that may prevent Microsoft Office applications from opening or functioning correctly.

**Location**

File → Options → Add-ins → COM Add-ins → Go

**Result**

Successfully reviewed installed COM Add-ins and identified where Office extensions can be enabled, disabled, or removed during troubleshooting.

**Screenshot**
<img width="1918" height="972" alt="Image 1" src="https://github.com/user-attachments/assets/cf6ea710-4e52-4293-89d4-f4920045ffd8" />

## Scenario 2 – Office Applications Stuck in Task Manager

**Purpose**

Identify and end Microsoft Office applications that remain running in the background and prevent the application from opening correctly.

**Location**

Task Manager → Processes

**Result**

Verified Microsoft Office applications running in the background and demonstrated where to end a stuck process before attempting additional troubleshooting.

**Screenshot**
<img width="1918" height="971" alt="Image 2" src="https://github.com/user-attachments/assets/37fe8972-0448-4a2d-92f1-fd987d6bb893" />

## Scenario 3 – Excel Open and Repair

**Purpose**

Use Excel's built-in Open and Repair feature to recover damaged or corrupted workbooks before considering more advanced recovery methods.

**Location**

Excel → Open → Browse → Open ▼ → Open and Repair

**Result**

Successfully accessed the Open and Repair feature used to recover corrupted Excel workbooks.

**Screenshot**
<img width="1918" height="1008" alt="Image 3" src="https://github.com/user-attachments/assets/39a1e7a7-ea3a-4cac-bca5-db39525f424c" />

## Scenario 4 – Microsoft 365 Repair

**Purpose**

Review the Microsoft 365 repair options available when Office applications continue to experience issues after basic troubleshooting has been completed.

**Location**

Control Panel → Programs and Features → Microsoft 365 → Change

**Result**

Successfully accessed the Microsoft 365 repair options and identified Quick Repair and Online Repair as escalation steps before considering a full reinstallation.

**Screenshot**
<img width="1918" height="970" alt="Image 4" src="https://github.com/user-attachments/assets/0225ef24-aa59-49d9-b702-b8197b5b244c" />

# Lab 2B – Common Outlook Issues

## Scenario

A user reports multiple Microsoft Outlook issues, including Outlook failing to open, slow performance, search not working, email delivery problems, and authentication prompts. As the Help Desk technician, I investigated each issue using common Microsoft 365 and Windows troubleshooting techniques to restore Outlook functionality.

---

## Purpose

The purpose of this lab was to troubleshoot common Microsoft Outlook problems by diagnosing profile corruption, repairing Outlook data files, rebuilding the Windows Search index, verifying account connectivity, and applying standard Help Desk troubleshooting procedures.

---

## Location

- Microsoft Outlook
- Microsoft 365 Apps
- Control Panel
- Mail (Microsoft Outlook)
- Windows Search Service
- Windows Credential Manager
- File Explorer
- ScanPST.exe

---

## Objectives

- Troubleshoot Outlook startup issues.
- Launch Outlook in Safe Mode.
- Create a new Outlook profile.
- Repair Outlook data files using ScanPST.
- Rebuild the Outlook search index.
- Restart the Windows Search service.
- Troubleshoot send and receive issues.
- Verify Work Offline mode.
- Manage Outlook credentials using Credential Manager.
- Apply common Outlook troubleshooting best practices.

---

# Outlook Won't Open

## Issue

Outlook failed to launch due to a corrupt Outlook profile, damaged OST file, or conflicting add-ins.

## Troubleshooting Steps

- Tested Outlook in Safe Mode using `outlook.exe /safe`.
- Created a new Outlook profile using **Control Panel → Mail → Show Profiles → Add**.
- Allowed Outlook to recreate a new OST file and resynchronize mailbox data from Microsoft 365.

## Result

Creating a new Outlook profile resolved profile corruption and restored Outlook functionality.

**Screenshot:**
*Insert screenshot showing Mail → Show Profiles → Add Profile.*

---

# Slow Outlook Performance

## Issue

Outlook was running slowly due to a large mailbox, excessive add-ins, or possible Outlook data file corruption.

## Troubleshooting Steps

- Reviewed installed Outlook add-ins.
- Confirmed Office repair as a troubleshooting option.
- Used **ScanPST.exe** to scan and repair Outlook Personal Storage Table (PST) files when corruption was suspected.

## Result

Repairing Outlook data files and removing unnecessary add-ins improved Outlook performance and stability.

**Screenshot:**
*Insert screenshot showing ScanPST.exe.*

---

# Outlook Search Not Working

## Issue

Users were unable to locate emails because the Windows Search index had become corrupted or incomplete.

## Troubleshooting Steps

- Navigated to **File → Options → Search → Indexing Options → Advanced → Rebuild**.
- Restarted the **Windows Search** service through Services when required.

## Result

Rebuilding the search index restored Outlook search functionality.

**Screenshot:**
*Insert screenshot showing Outlook Indexing Options (or Windows Search Service).*

---

# Cannot Send or Receive Emails

## Issue

Users were unable to send or receive emails due to connectivity issues, incorrect configuration, Work Offline mode, antivirus interference, or Outlook rules.

## Troubleshooting Steps

- Verified internet connectivity.
- Confirmed account configuration.
- Checked Outlook Send/Receive settings.
- Verified Work Offline mode was disabled.
- Reviewed antivirus software and Outlook rules that could block mail flow.

## Result

After correcting the Outlook configuration and disabling Work Offline mode, email functionality was restored.

**Screenshot:**
*Insert screenshot showing the Outlook Send/Receive ribbon with Work Offline.*

---

# Frequent Outlook Crashes

## Issue

Outlook frequently crashed because of conflicting add-ins, outdated Office applications, or profile corruption.

## Troubleshooting Steps

- Reviewed installed COM Add-ins.
- Confirmed Microsoft Office was fully updated.
- Recreated the Outlook profile when profile corruption was suspected.

## Result

Removing conflicting add-ins and rebuilding the Outlook profile resolved repeated Outlook crashes.

---

# Password Prompt Keeps Appearing

## Issue

Users were repeatedly prompted for their Microsoft 365 password due to cached credential corruption or authentication issues.

## Troubleshooting Steps

- Verified account credentials.
- Opened **Control Panel → Credential Manager → Windows Credentials**.
- Removed outdated cached Outlook credentials and allowed Outlook to authenticate again.

## Result

Refreshing cached credentials eliminated repeated authentication prompts.

**Screenshot:**
*Insert screenshot showing Windows Credential Manager.*

---

# Attachments Won't Open

## Issue

Users were unable to open email attachments because of blocked file types, incorrect file associations, or Outlook add-in conflicts.

## Troubleshooting Steps

- Verified Outlook security settings.
- Checked Windows default file associations.
- Reviewed Outlook add-ins for conflicts.

## Result

Correcting file associations and resolving Outlook configuration issues restored attachment functionality.

---

# Outlook Data File (.PST) Issues

## Issue

A corrupt Outlook PST file caused mailbox errors and Outlook instability.

## Troubleshooting Steps

- Used **ScanPST.exe** to scan and repair the Outlook data file.
- Confirmed mailbox integrity after the repair completed successfully.

## Result

Repairing the PST file restored Outlook stability and resolved mailbox corruption.

---

# Calendar Issues

## Issue

Shared calendars were not synchronizing correctly because of profile corruption, synchronization problems, or insufficient permissions.

## Troubleshooting Steps

- Verified Outlook synchronization settings.
- Recreated the Outlook profile when required.
- Confirmed calendar permissions for shared mailboxes.

## Result

Calendar synchronization resumed after correcting profile and permission issues.

---

# Outlook Rules

## Issue

Incorrect Outlook rules automatically moved or deleted incoming emails.

## Troubleshooting Steps

- Reviewed configured Outlook rules.
- Compared mailbox behaviour using Outlook on the Web (OWA).
- Confirmed whether the issue was client-side or mailbox-related.

## Result

Removing incorrect mailbox rules restored expected email delivery behaviour.

---

## Skills Demonstrated

- Microsoft Outlook Troubleshooting
- Microsoft 365 Administration
- Outlook Profile Management
- Outlook Safe Mode
- ScanPST Inbox Repair Tool
- Windows Search Indexing
- Windows Services
- Windows Credential Manager
- Mail Flow Troubleshooting
- Outlook Rules
- Root Cause Analysis
- Help Desk Troubleshooting
- End-User Support
- Microsoft Windows Administration
