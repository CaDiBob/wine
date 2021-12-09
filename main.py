import datetime
import os
import pandas
import collections
from dotenv import load_dotenv
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_age_winery():
    foundation_year = 1920
    this_year = datetime.datetime.now()
    return this_year.year - foundation_year


def get_wines_assortment(data_base):
    wines_data = pandas.read_excel(
                    data_base,
                    sheet_name='Лист1',
                    na_values=False,
                    keep_default_na=False
                    ).to_dict(orient='records')
    categories = collections.defaultdict(list)
    for wine in wines_data:
        categories[wine['Категория']].append(wine)
    return categories


def render_page(categories, env, date):
    template = env.get_template('template.html')
    rendered_page = template.render(
                    categories=categories,
                    age_winery=f'Уже {date} год с вами'
                    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


def main():
    load_dotenv()
    data_base = os.getenv('DB_XLSX', default='wine.xlsx')
    categories = get_wines_assortment(data_base)
    date = get_age_winery()
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    render_page(categories, env, date)


if __name__ == '__main__':
    main()
