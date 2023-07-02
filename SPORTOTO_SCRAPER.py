import requests
import xlsxwriter
import urllib3
import datetime as dt
import time as t


def write_to_sheet1(row_index, gameroundName, date, time, ilkyari_skor, mac_skor, full_time_win, home_team_name, away_team_name, workbook, sheet):

    bold_format = workbook.add_format({'bold': True})
    center_format = workbook.add_format({'align': 'center'})
    bg_color_format = workbook.add_format({'bg_color': 'black', 'font_color': 'white'})
    title_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'black', 'font_color': 'white'})
    center_bold_format = workbook.add_format({'bold': True, 'align': 'center'})

    column_widths = {
        'A': 27.25,
        'B': 15,
        'C': 12,
        'D': 10,
        'E': 10,
        'F': 10,
        'G': 27.25,
        'H': 27.25,
    }

    for column, width in column_widths.items():
        sheet.set_column(column + ':' + column, width)

    sheet.set_row(0, None, bg_color_format)
    sheet.set_column('A:H', None, bold_format)
    sheet.set_row(0, None, center_format)

    # A sütunu için ayrıca genişlik ayarı yap
    sheet.set_column('A:A', 27.25)

    sheet.write('A1', 'Sezon', title_format)
    sheet.write('B1', 'Tarih', title_format)
    sheet.write('C1', 'Saat', title_format)
    sheet.write('D1', 'İY', title_format)
    sheet.write('E1', 'MS', title_format)
    sheet.write('F1', 'Toto', title_format)
    sheet.write('G1', 'Ev Sahibi', title_format)
    sheet.write('H1', 'Deplasman', title_format)

    sheet.write(row_index, 0, gameroundName or "-", center_bold_format)
    sheet.write(row_index, 1, date or "-", center_bold_format)
    sheet.write(row_index, 2, time or "-", center_bold_format)
    sheet.write(row_index, 3, ilkyari_skor or "-", center_bold_format)
    sheet.write(row_index, 4, mac_skor or "-", center_bold_format)
    sheet.write(row_index, 5, full_time_win, center_bold_format)
    sheet.write(row_index, 6, home_team_name or "-", center_bold_format)
    sheet.write(row_index, 7, away_team_name or "-", center_bold_format)

    sheet.autofilter('A1:H1')

    return row_index + 1


def write_to_sheet2(row_index2, gameroundName, fifteenWinCount, fifteenWinPrize, fourteenWinCount, fourteenWinPrize, thirteenWinCount, thirteenWinPrize, twelveWinCount, twelveWinPrize, workbook, sheet2):
    bold_format = workbook.add_format({'bold': True})
    center_format = workbook.add_format({'align': 'center'})
    bg_color_format = workbook.add_format({'bg_color': 'black', 'font_color': 'white'})
    title_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'black', 'font_color': 'white'})
    center_bold_format = workbook.add_format({'bold': True, 'align': 'center'})

    column_widths = {
        'A': 27.25,
        'B': 12,
        'C': 18,
        'D': 12,
        'E': 18,
        'F': 12,
        'G': 18,
        'H': 12,
        'I': 18,
    }

    for column, width in column_widths.items():
        sheet2.set_column(column + ':' + column, width)

    sheet2.set_row(0, None, bg_color_format)
    sheet2.set_column('A:I', None, bold_format)
    sheet2.set_row(0, None, center_format)

    sheet2.set_column('A:A', 27.25)

    sheet2.write('A1', 'Sezon|Hafta', title_format)
    sheet2.write('B1', '15 Bilen', title_format)
    sheet2.write('C1', 'İkramiye', title_format)
    sheet2.write('D1', '14 Bilen', title_format)
    sheet2.write('E1', 'İkramiye', title_format)
    sheet2.write('F1', '13 Bilen', title_format)
    sheet2.write('G1', 'İkramiye', title_format)
    sheet2.write('H1', '12 Bilen', title_format)
    sheet2.write('I1', 'İkramiye', title_format)

    sheet2.write(row_index2, 0, gameroundName or "-", center_bold_format)
    sheet2.write(row_index2, 1, fifteenWinCount or "-", center_bold_format)
    sheet2.write(row_index2, 2, fifteenWinPrize or "-", center_bold_format)
    sheet2.write(row_index2, 3, fourteenWinCount or "-", center_bold_format)
    sheet2.write(row_index2, 4, fourteenWinPrize or "-", center_bold_format)
    sheet2.write(row_index2, 5, thirteenWinCount or "-", center_bold_format)
    sheet2.write(row_index2, 6, thirteenWinPrize or "-", center_bold_format)
    sheet2.write(row_index2, 7, twelveWinCount or "-", center_bold_format)
    sheet2.write(row_index2, 8, twelveWinPrize or "-", center_bold_format)

    sheet2.autofilter('A1:I1')

    return row_index2 + 1

def main():

    import time
    
    row_index = 1
    row_index2 = 1
    
    URL_TOTO = "https://webapi.sportoto.gov.tr/api/GameMatch/GetGameMatches/?"
    URL_RESULT_TOTO = "https://webapi.sportoto.gov.tr/api/GameResult/GetGameResultByGameRoundId?"


    header = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36"
    }

    tarih = dt.datetime.now().strftime("%d_%m_%Y")
    dosya_adı = f"{tarih}_guncel.xlsx"
    
    workbook = xlsxwriter.Workbook(dosya_adı)
    sheet = workbook.add_worksheet('Sheet1')
    sheet2 = workbook.add_worksheet('Sheet2')

    menu = input('''
(1) Belirtilen Haftanın Sportoto Sonucunu Getir ve Excele Aktar.
(2) Tüm Sezona Ait Sportoto Sonuçlarını Çek ve Excele Atkar.\n >>> ''')
    
    if menu == "1":
        URL_CHECK_ID = int(input("Hangi Haftanın Sonucunu Çekmek İstiyorsunuz? : "))
        id_check = 1
    else:
        URL_CHECK_ID = 300 # başlangıç haftası
        id_check = 48 # 48 hafta var şuan bu bilgi değişir.

    while True:
        for URL_CHECK_ID in range(URL_CHECK_ID,  URL_CHECK_ID + id_check):
            try:
                
                params = {
                    "gameRoundId": URL_CHECK_ID
                }

                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                response = requests.get(URL_TOTO, params=params, headers=header, verify=False)
                response = response.json()

                if response is None:
                    print("Yanıt alınamadı.")
                    continue

                if 'isSucceed' in response:
                    matchs = response['message']
                    game_round_name = response['object'][0]['gameRoundName']
                    print(game_round_name, " - ",matchs)
                    print("-:-"*55)

                    matches = response['object']  # Tüm "match" öğelerini içeren liste

                    gameroundName = response['object'][0]['gameRoundName']

                    for match in matches:

                        match_details = match['match']
                        
                        gameroundName = response['object'][0]['gameRoundName']
                        
                        datetime_str = match_details['date']
                        datetime_obj = dt.datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")
                        date = datetime_obj.strftime("%d-%m-%Y")
                        time = datetime_obj.strftime("%H:%M:%S")
                        
                        home_team_name = match_details['homeTeam']['name'] if match_details.get('homeTeam') is not None and match_details['homeTeam'].get('name') is not None else '-'
                        away_team_name = match_details['awayTeam']['name'] if match_details.get('awayTeam') is not None and match_details['awayTeam'].get('name') is not None else '-'

                        home_halftime_score = match_details['score']['homeHalfTime'] if match_details.get('score') and 'homeHalfTime' in match_details['score'] else 0
                        away_halftime_score = match_details['score']['awayHalfTime'] if match_details.get('score') and 'awayHalfTime' in match_details['score'] else 0

                        home_fulltime_score = match_details['score']['homeRegular'] if match_details.get('score') and match_details['score'].get('homeRegular') else 0
                        away_fulltime_score = match_details['score']['awayRegular'] if match_details.get('score') and match_details['score'].get('awayRegular') else 0

                        if home_fulltime_score == away_fulltime_score:
                            full_time_win = 0
                        else:
                            full_time_win = match_details['fullTimeWin']

                        ilkyari_skor = "-".join([str(home_halftime_score), str(away_halftime_score)])
                        mac_skor = "-".join([str(home_fulltime_score), str(away_fulltime_score)])

                        result = f"{gameroundName} | {date} | {time} | {ilkyari_skor} | {mac_skor} | {full_time_win} | {home_team_name} - {away_team_name}"

                        print(result)

                        row_index = write_to_sheet1(
                            row_index,
                            gameroundName,
                            date,
                            time,
                            ilkyari_skor,
                            mac_skor,
                            full_time_win,
                            home_team_name,
                            away_team_name, 
                            workbook,
                            sheet
                        )


                params = {
                    "id": URL_CHECK_ID
                }

                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                response = requests.get(URL_RESULT_TOTO, params=params, headers=header, verify=False)
                response = response.json()

                if 'isSucceed' in response:
                    results = response['message']
                    gameRoundClose = response['object']['gameRoundCloseDate']

                    datetime_str = gameRoundClose
                    datetime_obj = dt.datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")
                    dates = datetime_obj.strftime("%d-%m-%Y")
                    times = datetime_obj.strftime("%H:%M:%S")

                    if response['object']['resultDescription'] == "Tebrikler":
                        fifteenWinCount = response['object']['fifteenWinCount']
                        fifteenWinPrize = response['object']['fifteenWinPrize']
                        fourteenWinCount = response['object']['fourteenWinCount']
                        fourteenWinPrize = response['object']['fourteenWinPrize']
                        thirteenWinCount = response['object']['thirteenWinCount']
                        thirteenWinPrize = response['object']['thirteenWinPrize']
                        twelveWinCount = response['object']['twelveWinCount']
                        twelveWinPrize = response['object']['twelveWinPrize']
                        
                    
                    print("-:-"*55)
                    print("Kapanış : ", dates,"-", times, " - ",results)
                    print("-"*55)
                    print("15 Bilen :",fifteenWinCount, "|Kazanç ", fifteenWinPrize," TL")
                    print("14 Bilen :",fourteenWinCount, "|Kazanç ", fourteenWinPrize," TL")
                    print("13 Bilen :",thirteenWinCount, "|Kazanç ", thirteenWinPrize," TL")
                    print("12 Bilen :",twelveWinCount, "|", twelveWinPrize," TL")
                    print("-:-"*55)
                    
                    row_index2 = write_to_sheet2(
                        row_index2,
                        gameroundName,
                        fifteenWinCount,
                        fifteenWinPrize,
                        fourteenWinCount,
                        fourteenWinPrize,
                        thirteenWinCount,
                        thirteenWinPrize,
                        twelveWinCount,
                        twelveWinPrize,
                        workbook,
                        sheet2,
                    )

                t.sleep(1) #atlatmak.
                
            except Exception as e:
                print("Bir hata oluştu:", str(e))
                print("Lütfen geçerli bir ID girin ve tekrar deneyin.")
                continue

        URL_CHECK_ID = URL_CHECK_ID + id_check
        workbook.close()

        choice = input("Devam etmek istiyor musunuz? (E/h): ")
        if choice.lower() != "e":
            print("İşlemleriniz Sonlandırıldı. Teşekkürler.")
            break
        
    
if __name__ == "__main__":
    main()
