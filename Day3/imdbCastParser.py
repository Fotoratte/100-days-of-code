from imdb import IMDb
import xlsxwriter
import argparse

COLUMN_WIDTH = 2


class SeriesResult:
    title = "Dummy-title"
    episode_cast_dict = dict()
    season_episode_dict = dict()

    def __init__(self, title, season_episode_dict, episode_cast_dict):
        self.title = title
        self.season_episode_dict = season_episode_dict
        self.episode_cast_dict = episode_cast_dict


def parse_args():
    parser = argparse.ArgumentParser(description='Creates xls-file from imdb.com')
    parser.add_argument('series', type=str, help='imdb.com-ID of the series')
    parser.add_argument('-s', '--search', required=False, action='store_true', help='Search for the series name')
    parser.add_argument('-o', '--output-file', required=False, help='name of the xls-file')

    return parser.parse_args()


def search_for_series(ia, series_name):
    series = ia.search_movie(title=series_name)
    for s in filter(lambda x: x['kind'] == 'tv series', series):
        print("ID: {}, Title: {}, Kind: {}".format(s.movieID, s['title'], s['kind']))


def parse_series(ia, series_id):
    series = ia.get_movie(series_id)

    ia.update(series, 'episodes')

    episode_cast_dict = dict()

    season_episode_dict = dict()

    for seasonNumber in sorted(series['episodes'].keys()):
        print("Parsing Season {} of {}".format(seasonNumber, len(series['episodes'].keys())))
        season_episode_dict[seasonNumber] = len(series['episodes'][seasonNumber])
        for episodeNumber in sorted(series['episodes'][seasonNumber]):
            episode = series['episodes'][seasonNumber][episodeNumber]
            ia.update(episode, 'full credits')
            for cast in episode['cast']:
                if cast not in episode_cast_dict:
                    episode_cast_dict[cast] = list()
                episode_cast_dict[cast].append('{}:{}'.format(str(seasonNumber), str(episodeNumber)))

    return SeriesResult(series['title'], season_episode_dict, episode_cast_dict)


def save_to_file(series_result):
    series_title = series_result.title
    print("Saving results into {}.xlsx".format(series_title))
    workbook = xlsxwriter.Workbook('{}.xlsx'.format(series_title))
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format()
    cell_format.set_bg_color('green')

    season_episode_dict = series_result.season_episode_dict
    row_number = 0
    column_number = 0
    worksheet.write(row_number, column_number, 'Character')
    column_number = column_number + 1;
    for seasonNumber in season_episode_dict.keys():
        worksheet.write(row_number, column_number, "S{}".format(str(seasonNumber)))
        worksheet.set_column(column_number, column_number + season_episode_dict[seasonNumber], COLUMN_WIDTH)
        column_number = column_number + season_episode_dict[seasonNumber]
    print(" ")

    episode_cast_dict = series_result.episode_cast_dict
    max_character_name = 0
    for cast, episodes in sorted(episode_cast_dict.items(), key=lambda x: len(x[1]), reverse=True):
        column_number = 0
        row_number = row_number + 1
        actor_name = "{} ({})\t".format(cast['name'], cast.currentRole)
        max_character_name = max(max_character_name, len(actor_name))
        worksheet.write(row_number, column_number, actor_name)
        for seasonNumber in season_episode_dict.keys():
            for episodeNumber in range(1, season_episode_dict[seasonNumber] + 1):
                column_number = column_number + 1
                if '{}:{}'.format(seasonNumber, episodeNumber) in episodes:
                    worksheet.write_blank(row_number, column_number, ' ', cell_format)
                else:
                    worksheet.write_blank(row_number, column_number, ' ')

    worksheet.set_column(0, 0, max_character_name)
    workbook.close()


def __main__():
    ia = IMDb()

    args = parse_args()

    if args.search:
        search_for_series(ia, args.series)
    else:
        series_result = parse_series(ia, args.series)

        if args.output_file is not None:
            series_result.title = args.output_file

        save_to_file(series_result)


__main__()
