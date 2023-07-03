# sportotoWebApi_dataScraper
An API-based data scraper that allows you to efficiently extract and organize current or historical match data into Excel spreadsheets.

[Sportoto Gov TR - WEB API - MAÇ SONUÇLARI](https://webapi.sportoto.gov.tr/api/GameMatch/GetGameMatches/?gameRoundId=300)

[Sportoto Gov TR - WEB API - İKRAMİYE SONUÇLARI](https://webapi.sportoto.gov.tr/api/GameResult/GetGameResultByGameRoundId?id=300)

`pip install urllib3`
`pip install xlsxwriter`


`python SPORTOTO_SCRAPER.py `

   ```bash
(1) Belirtilen Haftanın Sportoto Sonucunu Getir ve Excele Aktar.
(2) Tüm Sezona Ait Sportoto Sonuçlarını Çek ve Excele Atkar.
 >>> 1
 Hangi Haftanın Sonucunu Çekmek İstiyorsunuz? : 300
