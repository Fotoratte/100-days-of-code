from imdb import IMDb


def __main__():
    ia = IMDb()

    series = ia.get_movie('3501584')

    ia.update(series, 'episodes')

    episode_cast_dict = dict()

    season_episode_dict = dict()

    for seasonNumber in sorted(series['episodes'].keys()):
        # for seasonNumber in range(1, 2):
        print("Season {}".format(seasonNumber))
        season_episode_dict[seasonNumber] = len(series['episodes'][seasonNumber])
        for episodeNumber in sorted(series['episodes'][seasonNumber]):
            episode = series['episodes'][seasonNumber][episodeNumber]
            ia.update(episode, 'full credits')
            for cast in episode['cast']:
                if cast not in episode_cast_dict:
                    episode_cast_dict[cast] = list()
                episode_cast_dict[cast].append('{}:{}'.format(str(seasonNumber), str(episodeNumber)))

    print("Character", end='\t')
    for seasonNumber in season_episode_dict.keys():
        print("S{}".format(str(seasonNumber)))
        for episodeNumber in range(1, season_episode_dict[seasonNumber]):
            print(" ", end='\t')
    print(" ")

    for cast, episodes in sorted(episode_cast_dict.items(), key=lambda x: len(x[1]), reverse=True):
        print("{} ({})\t".format(cast['name'], cast.currentRole), end='')
        for seasonNumber in season_episode_dict.keys():
            for episodeNumber in range(1, season_episode_dict[seasonNumber]):
                if '{}:{}'.format(seasonNumber, episodeNumber) in episode_cast_dict[cast]:
                    print(1, end='\t')
                else:
                    print(0, end='\t')
        print("")


__main__()
