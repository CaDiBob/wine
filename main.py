import datetime
import pandas
import collections
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def started_date():
    started_year = 1920
    now_year = datetime.datetime.now()
    return now_year.year - started_year


def reading_db():
    excel_data_df = pandas.read_excel(
                    'wine.xlsx',
                    sheet_name='Лист1',
                    na_values=False,
                    keep_default_na=False
                    )
    wines = excel_data_df.to_dict(orient='records')
    categories = collections.defaultdict(list)
    for wine in wines:
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
    categories = reading_db()
    date = started_date()
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    render_page(categories, env, date)


if __name__ == '__main__':
    main()
