import requests
from bs4 import BeautifulSoup
import pandas as pd
import time


class MovieParser:
    def __init__(self):
        self.base_url = "https://www.imdb.com"
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }

    def get_top_movies(self, limit=10):
        url = f"{self.base_url}/chart/top/"

        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, 'html.parser')
            movies = []

            movie_elements = soup.select('li.ipc-metadata-list-summary-item')[:limit]

            for movie in movie_elements:
                movie_data = self._parse_movie_element(movie)
                if movie_data:
                    movies.append(movie_data)
                time.sleep(0.1)

            return movies

        except requests.RequestException:
            return []

    def _parse_movie_element(self, movie_element):
        movie_data = {}

        title_elem = movie_element.select_one('h3.ipc-title__text')
        if title_elem:
            movie_data['title'] = title_elem.text.strip()

        rating_elem = movie_element.select_one('span.ipc-rating-star--rating')
        if rating_elem:
            movie_data['rating'] = float(rating_elem.text.strip())

        year_elem = movie_element.select_one('span.cli-title-metadata-item')
        if year_elem:
            movie_data['year'] = year_elem.text.strip()

        return movie_data

    def save_to_excel(self, movies, filename='movies.xlsx'):
        if not movies:
            return

        df = pd.DataFrame(movies)
        df = df.sort_values('rating', ascending=False)
        df = df[['title', 'year', 'rating']]

        df.to_excel(filename, index=False)


def main():
    parser = MovieParser()
    movies = parser.get_top_movies(limit=10)

    if movies:
        parser.save_to_excel(movies, 'movies.xlsx')


if __name__ == "__main__":
    main()