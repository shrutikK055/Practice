import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

url = "https://www.imdb.com/chart/top/?ref_=nv_mv_250"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    soup = BeautifulSoup(response.content, "html.parser")
    
    # Locate the divs with movie information
    rows = soup.select("td.titleColumn")

    if not rows:
        print("❌ Could not find the movie rows. HTML structure may have changed.")
    else:
        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Top 100 IMDb Movies"
        ws.append(["Rank", "Title", "Year", "Rating", "Link"])

        for idx, row in enumerate(rows[:100]):
            title = row.find("a").text.strip()
            year = row.find("span").text.strip("()")
            rating = row.find_next("td", class_="imdbRating").strong.text.strip()
            link = "https://www.imdb.com" + row.find("a")["href"]

            ws.append([idx + 1, title, year, rating, link])
            print(f"{idx + 1}. {title} ({year}) - Rating: {rating}")

        wb.save("IMDb_Top_100.xlsx")
        print("\n✅ IMDb Top 100 saved to 'IMDb_Top_100.xlsx'")

else:
    print(f"❌ Request failed. Status code: {response.status_code}")
