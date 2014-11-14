OkCupidAutoBot v. 1.0

Purpose: Scrapes OkCupid data and votes on profiles (score 3-5)

	1) By visiting and voting on profiles of your prospects,
	   your profile will be placed at the top of their search results.
	2) The program has the option to write to DB or Excel sheet.

Sql server Setup
---

Create a database OKCUPIDDB01 and run these two files.
	1) CREATE_TABLE_Dim_Profile.sql
	2) CREATE_TABLE_Fact_Essay.sql

Configuration
OkCupidAutoBot.exe.config
---

1) username
	Your OkCupid username.

2) password
	Your OkCupid password.

3) matchThreshhold
	The minimum match percentage at which you'll scrape a profile.

4) sqlConnection
	SQL Connection string
	Example: Data Source=SERVER_NAME;Initial Catalog=DATABASE_NAME;Integrated Security=True

5) profileFields
	Profile dimension fields (ex. username, age, etc.)

6) essayFields
	Essay fact fields, used for storing user's essays (Self-Summary, Six things, etc.)

7) sqlOrExcel
	Choose to write to db or excel sheet.
	* The path to the exce sheet is /AssEMBLY_ROOT_DIR/excel/girls.xlsx

8) detailsTable
	Name of the profile details SQL table.

9) essaysTable
	Name of the essays SQL table.

10) baseUri
	Base URI to OkCupid.  For ex: https://www.okcupid.com

11) girlsPerSession
	The number of profiles to visit and process per session.

12) matchesQueryString
	The query string used to download matches.
	See: http://www.reddit.com/r/OkCupid/comments/qi8iw/understanding_okc_url_manipulation/

13) callThrottling
	Will place a limit on the number of calls made per second.

