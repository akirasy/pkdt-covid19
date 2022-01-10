# COVID19 Management System

## Introduction

Once a patient is diagnosed as positive COVID19, Health Official should take several actions in
containing the outbreak. Healthcare worker will investigate the patient. Soon after, healthcare
worker required to do contact tracing, patient report and patient referral.
In overcoming the difficulties of writing report and miscommunication, this app is born.

## Template links
1. [Clerking Template (GoogleDoc)](https://docs.google.com/document/d/13yhrGK3GNjTIWQgIKfG75Ad0uET35qNaMBaeJo6CVzY/edit?usp=sharing)
1. [Request Access Form (GoogleForm)](https://docs.google.com/forms/d/1g1p6KtD7P-F4tVAVBWuFSvu9mf9leQTVvpdXZWaapcI/edit?usp=sharing)
1. [Kes Positif COVID19 (GoogleSpreadsheet)](https://docs.google.com/spreadsheets/d/1MQvczLh6cmX5DH9uKTi6w0zM_Xr_QVAoKYZACTShv9E/edit?usp=sharing)

## Pre-Installation

1. Setup GoogleDrive file directory.
    ```
    parent_folder
      ├── generated-file
      │     ├── current
      │     └── archive
      │           ├── Year 2021
      │           └── Year 2022
      └── workdir
    ```
1. Copy template GoogleForm for user registration.
    - Move the GoogleForm to `workdir` folder.
    - Link GoogleForm output to new GoogleSpreadsheet.
    - Rename `Form Responses 1` to `User Access List`.
    - Move the GoogleSpreadsheet to `workdir` folder.
    - (Optional) Rename GoogleSpreadsheet to `User Access List`.

1. Copy template GoogleSpreadsheet (Kes Positif COVID19).
    - Move the GoogleSpreadsheet to `workdir` folder.
    - (Optional) Rename GoogleSpreadsheet to `Kes Positif COVID19`.

1. Copy template GoogleDoc (Clerking Template)
    - Move the GoogleDoc to `workdir` folder.
    - (Optional) Rename GoogleDoc to `Clerking Template`.

1. You will end up with file directory like this.
    ```
    parent-folder
      ├── generated-file
      │     ├── current
      │     └── archive
      │           ├── Year 2021
      │           └── Year 2022
      └── workdir
            ├── Clerking Template (GoogleDoc)
            ├── Request Access Form (GoogleForm)
            ├── User Access List (GoogleSpreadsheet)
            └── Kes Positif COVID19 (GoogleSpreadsheet)

    ```

## Installation

### 1. Enable AppScript on GoogleSpreadsheet
1. Open up Kes Positif COVID19 (GoogleSpreadsheet).
1. At the menu bar, open up `Extensions -> AppScript`.
1. A new tab will open and shows `AppScript Editor`.
1. Copy all code [here](src/appScript/) to the `AppScript Editor`.
1. Simply save and close tab.

### 2. Edit corresponding value
1. Open up Kes Positif COVID19 (GoogleSpreadsheet).
1. Go to tab `appScript.gs`.
1. Edit these parameter.
    ```
    - `clerking_template`  # File ID for Clerking Template (GoogleDoc)
    - `TLH_folder_main`    # Folder ID for `parent-folder/generated-file/current`
    - `TLH_prefix`         # Prefix reg number
    - `spreadsheet_owner`  # Email
    - `spreadsheet_uac_id` # File ID for User Access List (GoogleSpreadsheet)
    ```

## 3. Set file permission
1. `Clerking Template (GoogleDoc)` -> With link can view
1. `User Access List (GoogleSpreadsheet)` -> Private
1. `Kes Positif COVID19 (GoogleSpreadsheet)` -> Private

## Usage
1. Each user will be notified that `This app is not verified by Google`.
2. Follow [this](src/enable-appscript.pdf) tutorial to use the AppScript.
> The installed app (AppScript) is not verified by Google and continuing to use this app is at user's own risk.

## Privacy Policy
1. User's data and credential are stored in `User Access List (GoogleSpreadsheet)` and should not be used for other purpose.
2. Patient's data are confidential and all user should maintain its privacy.

## License

Copyright (C) 2021 akirasy
```
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.
```
