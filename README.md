# Master Case Order App

- Copy Orders to Suppliers' Sheet
- Download Images for Selected Orders
- Pull Image URLs From Google Drive

(2)
SZ5 & EMB2 added col I,M,Q,U,Y for image links!!!
can you update script for "copy&paste"

(1)
https://docs.google.com/spreadsheets/d/1N54zulgyi8D2uR_cvsUcQhktRSwfU0iTsec7J4N1fMI/edit?usp=sharing
[DOWNLOAD]

- select row(s) click download, if filter, dun download hidden rows
- after download, change row to green
- if filtered, hidden row dun download

[update tracking]

- if value in col E update (value to empty cell/change value in cell), copy and paste value to MASTER col F (please find from SHOP+ORDER NO.)
- - if value in col E deleted (with value -> empty cell), also delete in MASTER col F

(3)
[UPDATE]
select cell in master, (click update button), find order in case print folder, with old value remaining in comment + change cell color to orange

updatable values:
ADDRESS
PHONE 1&2&3&4&5
MATERIAL1&2&3&4&5
DESIGN 1&2&3&4&5 (select design no. -> update image link& preview!!)

[SEARCH FOR DUPLICATES]
-seach for same recipient on address (per shop!) when rows filtered!!!!

(4)
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

(5) cancel image update trigger
