# Google Sheets / Forms NFL Pick'Ems and Survivor Pool, v1.2.0
## Creation and Management Tool for Running your Own Group

Google Sheet document with multiple script files to generate Google Forms for season-long NFL Pick'Ems or Survivor league management

-------------------------

**TLDR: Go [here](https://docs.google.com/spreadsheets/d/17k_Rj8QwprbY07odPzRSmhHj9d0e7f6s532YC5Jw8F8/edit?usp=sharing) and make a copy of the sheet. Follow prompts. Enjoy!**

-------------------------


**Welcome!** The project below was developed over four seasons of NFL play (maybe NCAAF eventually) to create a new way of managing an NFL pick ’ems, survivor, or eliminator pool. Copying the file will enable you to customize and generate a series of sheets in your copy of the spreadsheet for tracking all picks through the 18 regular season games of an NFL season and may work through the playoffs. It also includes a Monday Night Football most correct season-long winner, a weekly most correct winner, and a season-long most correct winner. The tool will also create a weekly Google Form (questionnaire) that is used to collect responses from members that can be imported to the spreadsheet easily. Match results and tiebreaker scores can also be pulled in via scripts in the 'Picks' menu. The final Monday Night Football game score total each week is used as the tiebreaker for the pick ‘ems weekly competition (some weeks we do have 2 MNF games). Tiebreakers, comments, exclusion of Thursday games, and more can be disabled/enabled via the setup.

It’s up to the person running the league to import the picks for the week (ideally before Thursday night) and also to update the form for the coming week (usually done Tuesday or Wednesday morning to send to the members).

I was keen to help a friend create a more robust way to track a family and friends league three seasons ago and the effort resulted in this massive and complex block of thousands of lines of code. I’m not a coder by training, I’m an industrial designer and product manager. I hope it doesn’t break for you--but let me know if it does! If you’re inclined and have enjoyed the script and care to support my wife, my five kiddos (sixth on the way!), and me, you can [buy me a coffee](https://www.buymeacoffee.com/benpowers)--no pressure though, I’m just excited that you’re using this tool!


**Disclaimer:** This set of functions relies on the use of the ESPN API for pulling NFL game data. You can find the ESPN terms of use [here](https://disneytermsofuse.com/). I’m sharing these scripts with you with the intent that you are taking on the responsibilities of the terms of use for your own personal use and don’t condone or endorse your use of the code here for monetization of “apps” or any other content. The terms outline the need for an “Information Form” to be submitted by a parent or guardian if you are a minor. This content is not intended to be published nor executed outside of the use by personal users. 

Lastly, there are some safeguards Google has in place to avoid allowing users to execute any malicious code from the Google Scripts console. Please feel free to review the code, as it contains no functions to share information, transfer information, or send emails. Information only travels between your personal Google Sheet and your Google Form (copied from a template form) that are created in the process outlined below. All sharing of content must be done by you directly (via the links that are created), such as sharing the link to the Google Form with your members and sharing a “view only” version of the spreadsheet with them to allow the members to see their league’s standings.

-------------------------

**Notable Changes**
This newest update to version 1.2.0 includes the following new improved features/changes:
- **More intuitive setup and initialization tool**
- Use of **Document Properties** (formerly used _Script_ Properties)
  - Storing timezone, initialization, configuration, members, and form details
- Use of a back-end Sheet ("database") to record pick selections from the form as a quicker means to recall user picks
- **Against the Spread** picks for Pick 'Ems/Survivor/Eliminator (by popular demand)
-   Fetches from API, can be automated
-   Allows for overwriting (hover over existing)
-   Provides ability to manually enter ATS values
-   Custom selection of matchups (by weekday or individually)
- Survivor & Eliminator (new):
  - "Lives" (1-3 for now)
  - "Revive" option
  - Custom start week (should allow for restarts)
- **Member Manager** Panel:
  - "Paid" status marking
  - Delete members
  - Drag-and-drop adjustment of order
- **Form Manager** Panel:
  - To copy form link, open form, or edit form
  - To "Lock" and "Unlock" forms
  - Enable "Auto-Sync"
  - Display form features
  - Preview form responses
  - Show new members who've joined
- **Triggers**:
  - To keep the Survivor/Eliminator pools correct (if being used)
  - To automatically fetch spreads (Tuesday-Saturday, provide a time)
  - To disable late form submissions (not tested yet)
- Your own **personal Form template** created to modify as you like, prompts you at first form creation
- **Emojis!** 🙂

-------------------------

## **Table of Contents**

### **1. Example Sheets** - Screenshots of the output from a league done in 2021 (Some of these have yet to be updated to visually represent the new format)

- NFL_OUTCOMES Sheet
- WEEKLY Sheet
- SUMMARY Sheet
- OVERALL Sheet
- MNF Sheet
- SURVIVOR Sheet
- ELIMINATOR Sheet
  
### **2. Example Form** - Screenshot of form from week 18 in 2021

### **3. Setup Instructions** - create new document, create script, paste code, run initial setup

### **4. Usage** - how to use the tool

### **5. Custom Functions Overview** - description of all custom functions in the “Picks” menu

-------------------------

# **1. Example Sheets (old versions shown, updates coming soon)** 

<h3 align="center">WEEKLY Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2024/08/2024_weekly_sheet.png" width="600" alt="WEEKLY Sheet">
</p>

<h3 align="center">NFL OUTCOMES Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/09/googlesheets-pickems-outcomes-sheet.png" width="600" alt="NFL OUTCOMES">
</p>

<h3 align="center">SUMMARY Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2024/08/2024_summary_sheet.png" width="600" alt="SUMMARY Sheet">
</p>

<h3 align="center">MNF Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example03.png" width="600" alt="MNF Sheet">
</p>

<h3 align="center">OVERALL Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example04.png" width="600" alt="OVERALL Sheet">
</p>

<h3 align="center">RANK Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2024/08/2024_rank_sheet.png" width="600" alt="RANK Sheet">
</p>

<h3 align="center">SURVIVOR Sheet</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example05.png" width="600" alt="SURVIVOR Sheet">
</p>

<h3 align="center">ELIMINATOR Sheet (showing survivor)</h3>
<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example05.png" width="600" alt="SURVIVOR Sheet">
</p>


-------------------------

# **2. Example Form (needs updated to 2026 version)**
Update your form to look like this, or whatever you prefer. The script will create all the weekly entries for each matchup of the week, a survivor pool prompt, a tiebreaker entry field, and a comments section. When membership is unlocked, the form will have a text entry field, rather than the dropdown, for “Name”.

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example06.png" width="500" alt="Example Form part 1">
</p>

<h3 align="center">[MANY MATCHES LATER]</h3>

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-example07.png" width="500" alt="Example Form part 2">
</p>


-------------------------

# **3. Setup Instructions**
1. Go to my Google Sheet and **create a copy,** → [click here to open the spreadsheet](https://docs.google.com/spreadsheets/d/1cafdoM2H5JDXqDxH58Unww7TAFRHr_kWZN5IdDN2-eY/edit?usp=sharing)

2. An onOpen trigger will welcome you and prompt for timezone confirmation.

3. Once you run the "Initialize" function, an “Authorization required” box will appear, **click “Review permissions”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions06.png" width="600" alt="Review Permissions">
</p>

4. **Select your preferred Google account** for managing the spreadsheet and form

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions07.png" width="400" alt="Select Google Account">
</p>

5. "App isn't verified" pops up, **click “Advanced” on bottom left**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions08.png" width="400" alt="Advanced verification">
</p>

6. **Click “Go to NFL Picks (unsafe)”** on bottom left

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions09.png" width="400" alt="Got to project (unsafe) prompt">
</p>

7. Review permissions, scroll down and **click “Allow”**

<p align="center">
<img src="https://benpowerscreative.com/wp-content/uploads/2023/03/googlesheets-picks-instructions10.png" width="400" alt="Allow script to run">
</p>

8. You should be able to now **re-run the "Initialize"** function. Once this is done, you can start by creating a "Configuration" via the menu, then further options should follow.

4. **use the "Picks" > "Forms" > "Form Builder" function to create your first week form.**

5. Most functions are self-explanatory, but please go to the **"Extensions" > "Apps Script" > "picks.gs"** where there are some other descriptions at the top


-------------------------

1. Weekly usage:
 - **Share the Form** with your group
 - **Import picks** (going to "Picks" > "Forms" > "Form Import" via the menu will allow you to review provided submissions and import if desired (ideally done before the Thursday night game, if present).
 - Through the weekend, as games are completed you should be able to run the "Picks" > "Utilities" > "Fetch Scores" function and **import game outcomes** via that method
 - Survivor/Eliminator Only: Alternatively, enter the game outcomes manually on the "NFL_OUTCOMES" sheet
 - Pick ‘Ems: Alternatively, enter the game outcomes manually across the bottom of the correct weekly sheet. Note: If using a tiebreaker (sum of the last MNF game score), be sure to enter it in the cell to the right of the final match column or the weekly winner won’t be declared!
 - Upon completing the week (usually after the MNF game), you can **run the “Form Builder” function again** and start the process over again for the next week
 - **Repeat**

-------------------------

Hopeful improvements for future versions:

- Google User confirmation (auto-detection for submissions, tied to email and therefore unique identifier for members)
- Reorganize member names alphabetically as an option
- Multiple entries per user
- Option to have user removed upon submission from Form to avoid duplication
- Confidence pick 'ems capability
- Opting out of survivor competition in the Form
- NCAA Football capability
- More metrics (suggestions welcome!)

-------------------------

Thanks for checking out the project and for making it to the end!

