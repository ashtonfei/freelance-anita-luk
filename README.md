# Master Case Order App

- Copy Orders to Suppliers' Sheet
- Download Images for Selected Orders
- Pull Image URLs From Google Drive

### Description

Update 1>
SZ5 & EMB2 added col I,M,Q,U,Y for image links!!!
can you update script for "copy&paste"

Update 2>
select cell in master, (click update button), find order in case print folder, with old value remaining in comment + change cell color to orange

updatable values:
ADDRESS
PHONE 1&2&3&4&5
MATERIAL1&2&3&4&5
DESIGN 1&2&3&4&5 (select design no. -> update image link& preview!!)

Rules:
If "Address" selected, update "Address" and add old value to notes, set background to orange;
If "Phone" selected, update "Phone" and add old value to notes, set background to orange;
if "Model" selected, update "Model" and add old value to notes, set background to orange;
if "Design" selected, update "Link", "Preview", and add old value to notes, set background to orange;

Update 3>
SEARCH FOR DUPLICATES
-seach for same recipient on address (per shop!) when rows filtered!!!!

Tracking # APIs>
https://docs.google.com/spreadsheets/d/1W0CGWgI5ZpHxv_aqu8dpN-YBDuqnlyeF9SubvDrmPM0/edit?usp=sharing

https://drive.google.com/drive/folders/1uhi5F3FZECo7chWpdfsvjUyMagPDPgX7?usp=sharing

[SEARCH] (tab:paste here)

- find SS (not in subfolders) for name (col B), find order ID (col A), copy info to there
  col C -> col I
  col D -> col J
  col E -> col K
  col F -> col L
  col G -> col M
  col H -> col O
  col J -> col N

[CREATE LABEL] (tab:USPS)
(CHECK IF SHIPSTATION API WORKS FIRST)
http://www.shipstation.com/developer-api/
apisupport@shipstation.com
http://www.sfcservice.com/api
-copy and paste ??? selected rows to (tab:USPS)
NEED TO CREATE NEW ROW!!!
-NEED TO WORK UNDER FILTERED, DO NOT COPY HIDDEN!!
