from imdb import IMDb
import xlsxwriter
import argparse

COLUMN_WIDTH = 2


# SERIES_ID = '0944947'


def __main__():
    parser = argparse.ArgumentParser(description='Creates xls-file from imdb.com')
    parser.add_argument('series', type=str, help='imdb.com-ID of the series')
    parser.add_argument('-s', '--search', required=False, action='store_true', help='Search for the series name')
    parser.add_argument('-o', '--output-file', required=False, help='name of the xls-file')

    ia = IMDb()

    args = parser.parse_args()
    if args.search:
        series = ia.search_movie(title=args.series)
        for s in filter(lambda x: x['kind'] == 'tv series', series):
            print("ID: {}, Title: {}, Kind: {}".format(s.movieID, s['title'], s['kind']))

    else:

        series_id = args.series

        series = ia.get_movie(series_id)

        series_title = args.output_file

        if series_title is None:
            series_title = series['title']

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

        print("Saving results into {}.xlsx".format(series_title))
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(series_title))
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format()
        cell_format.set_bg_color('green')

        row_number = 0
        column_number = 0
        worksheet.write(row_number, column_number, 'Character')
        column_number = column_number + 1;
        for seasonNumber in season_episode_dict.keys():
            worksheet.write(row_number, column_number, "S{}".format(str(seasonNumber)))
            worksheet.set_column(column_number, column_number + season_episode_dict[seasonNumber], COLUMN_WIDTH)
            column_number = column_number + season_episode_dict[seasonNumber]
        print(" ")

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


__main__()
