# Changelog

Notable changes to the script, newest first, starting with version 1.2.0. Versions match the `VERSION` value at the top of `picks.gs`, which is also shown in the corner of the "Picks" > "Help & Support" window.

## 1.2.6 - 2026-09-17
- New MATRIX sheet that combines the TOTAL, PCT, RNK and MNF metrics into a single heatmap, deployed from "Matrix (TOT/RNK/PCT/MNF)" in the "Deploy / Refresh Sheets" menu
- After the MATRIX sheet is deployed you are offered the option to hide the TOTAL, PCT, RNK and MNF sheets it replaces, with a choice to not be asked again
- The TOTAL sheet's averages row now shows the average number of correct picks instead of formatting that average as a percentage

## 1.2.5 - 2026-09-16
- New "Update Form Template ID" and "Update Database ID" items in the Utilities menu let you paste a file ID or a full URL, review a report of the file's name, owner, type, trash status, and direct link, then confirm before it is saved
- Form help text now reads "Eliminated from Survivor" or "Eliminated from Eliminator" for a member with no lives left, and lost lives show as a red X instead of a black dot

## 1.2.4 - 2026-09-06
- After a pick import the Leaderboard switches to the week you just imported
- After a pick import, if the Leaderboard, Season, and Summary sheets are not all built yet, you are asked whether to deploy all the tracking sheets now, be reminded next week, or never be asked again
- Importing picks no longer fails in a week that has a Monday night game when MNF tracking is turned off
- Adding a member mid-week no longer breaks the pick import. The weekly sheet rebuild that runs when a new name appears used to stop partway, which left the picks grid blank and dropped that week's tiebreakers, comments, outcomes, margins, and spreads
- The Leaderboard's season percent correct column now fills in, instead of staying blank because the SUMMARY sheet saved that named range under the wrong name

## 1.2.3 - 2026-09-03
- New SEASON sheet showing each member's season total, percent correct, overall rank, a rank ticker for the move since last week, and a wide bar of every week colored by where that member finished, deployed from the new "Season Overview" item in the Deploy / Refresh Sheets menu
- New "Rebuild Survivor Sheet" and "Rebuild Eliminator Sheet" items in the Utilities menu that rebuild the sheet layout, restore every pick from the database, and re-grade lives across all weeks that have forms
- When tiebreakers are enabled, the weekly rank column now breaks ties with the tiebreaker once the actual tiebreaker score is entered, so the tiebreaker winner gets rank 1 and the runner-up gets rank 2 instead of both sharing a rank
- The Winner row on the weekly sheet is now filled with each winning team's colors. The Leaderboard weekday row above the matchups is now colored by day of the week. The Leaderboard WINNER row gets team colors too, but only the straight-up row, which ATS pools keep hidden
- The tracking sheet submenu is now called "Deploy / Refresh Sheets", its items have shorter names, and the whole submenu is hidden for pools that do not run Pick 'Ems
- The Survivor/Eliminator panel column headed "St" is now headed "Revived"

## 1.2.2 - 2026-08-19
- The Picks menu has a new "Tracking Sheets" submenu that deploys or refreshes any one sheet on its own (Leaderboard, Summary, Winners, TOT, RNK, PCT, MNF, Survivor, Eliminator, Contrarian, Pick Counts), or all of them at once
- A LEADERBOARD sheet is now built along with the rest of the tracking sheets: pick a week and a sort order in column A to see season totals, Survivor and Eliminator lives and status, and each member's picks, points, rank, percent, chances and tiebreaker for that week
- Two new sheets are available: CONTRARIAN shows how often each member picked against the group each week, and COUNTS tallies how many times each member picked each of the 32 teams and lists their most-picked teams
- You can now choose which game is the tiebreaker instead of always using the last game of the week, with a Tiebreaker column in the Form Builder, and both the form's tiebreaker question and the score fetched from the API follow that choice
- The Form Builder now shows whether a form already exists for the week you have selected, with options to open it, edit it, or copy its link, and weeks with no form yet are highlighted in the week list
- Importing picks no longer skips members who joined after the weekly sheet was created: rows are added for them and the picks already on the sheet are kept
- On the Survivor and Eliminator sheets, the lives, revives and status values were written into the wrong columns: the lives markers appeared under STATUS, the revive counts under LIVES, and each member's IN or OUT status under REVIVES. Those columns now line up with their headers, and rebuilding either sheet keeps the picks already recorded

## 1.2.1 - 2026-08-17
- A Survivor/Eliminator Manager panel now opens from the Picks menu, where you pick a member, switch between the Survivor and Eliminator pools, and see all 18 regular season weeks of picks, results, and lives in one table, with the option to change a pick, edit a life count, or revive a member for a chosen week
- The Spread Auto-Fetch panel is now included, so you can choose a day from Tuesday through Saturday and an hour for the weekly spread fetch, or turn the schedule off from the same panel
- Picks still waiting on a result now show in yellow on the SURVIVOR and ELIMINATOR sheets, each member gets an IN or OUT [WK#] status on a green or red background instead of a blank or a week number, a row under the pick grid counts how many members got each week right, and the SUMMARY sheet shows that same IN or OUT [WK#] status under a new STATUS heading
- A Survivor or Eliminator pick on a game with no winner entered yet is held as pending instead of costing a life, a tie now counts as a correct Survivor pick in a straight up pool, and the pool is no longer declared complete while games are still open
- Reviving a member now restores their lives for the weeks after the revive as well, and each revive is recorded against the week it was used
- Game data now comes from a different ESPN address, with the aim of avoiding denied and empty responses

## 1.2.0 - 2026-08-05
- The Utilities menu with Update NFL Data and Update Spread Data is now available before you create your first form
- The README has step by step instructions for routing the ESPN calls through a free Cloudflare Worker if you hit frequent API failures
- The overwrite prompt for spreads and over/unders now lists the games that changed, with your current numbers next to the incoming ones, up to five games plus a count of any others, and it only appears when something actually differs
- Updating NFL data now pulls the week you are importing instead of whatever week ESPN happens to be showing, and matches each game to your schedule by matchup, so preseason and playoff refreshes no longer drop games into the wrong week
- Spreads and over/unders you entered by hand are kept when ESPN has no line for a game, instead of being blanked out
