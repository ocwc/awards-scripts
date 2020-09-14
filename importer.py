import xlrd
import requests
from requests.auth import HTTPBasicAuth

import env


class Main(object):
    categories = []
    years = []
    items = []
    ids = {}

    images = {
        "2011": 12,
        "2012": 13,
        "2013": 14,
        "2014": 15,
        "2015": 16,
        "2016": 17,
        "2017": 18,
        "2018": 19,
        "2019": 20,
        "2020": 21,
    }

    def __init__(self):
        self._read_excel()
        self._create_taxonomies()
        self._create_posts()

    def _read_excel(self):
        wb = xlrd.open_workbook("data.xlsx")
        sheet = wb.sheet_by_name("AWARD WINNERS")
        for row_idx in range(2, sheet.nrows):
            data = {
                "year": str(int(sheet.cell(row_idx, 1).value)),
                "category": sheet.cell(row_idx, 2).value.strip(),
                "title": "{} {}".format(
                    sheet.cell(row_idx, 3).value, sheet.cell(row_idx, 4).value
                ).strip(),
                "institution": sheet.cell(row_idx, 5).value.strip(),
                "country": sheet.cell(row_idx, 10).value.strip(),
            }
            self.years.append(str(data["year"]))
            self.categories.append(data["category"])
            self.items.append(data)

        self.years = list(set(self.years))
        self.categories = list(set(self.categories))
        self.years.sort()
        self.categories.sort()

    def _create_taxonomies(self):
        for year in self.years:
            response = requests.post(
                "{}/wp/v2/award_year".format(env.HOST),
                auth=HTTPBasicAuth(env.USER, env.PASS),
                data={"name": year},
            )
            self.ids[year] = response.json()["id"]

        for category in self.categories:
            response = requests.post(
                "{}/wp/v2/award_category".format(env.HOST),
                auth=HTTPBasicAuth(env.USER, env.PASS),
                data={"name": category},
            )
            self.ids[category] = response.json()["id"]

    def _create_posts(self):
        for item in self.items:
            year = item["year"]
            requests.post(
                "{}/wp/v2/award".format(env.HOST),
                auth=HTTPBasicAuth(env.USER, env.PASS),
                data={
                    "title": item["title"],
                    "status": "publish",
                    "featured_media": self.images[year],
                    "award_year": self.ids[year],
                    "award_category": self.ids[item["category"]],
                    "institution": item["institution"],
                    "country": item["country"],
                },
            )


if __name__ == "__main__":
    Main()
