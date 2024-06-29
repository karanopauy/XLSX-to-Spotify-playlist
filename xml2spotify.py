import pandas as pd
import spotipy
from spotipy.oauth2 import SpotifyOAuth

def create_spotify_playlist_from_excel(file_path, access_token):
    # Authenticate with Spotify using the access token
    sp = spotipy.Spotify(auth=access_token)

    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Extract song names and artist names from the DataFrame
    songs = df.iloc[:, 0]
    artists = df.iloc[:, 1]

    # Create a new playlist on Spotify
    playlist_name = input("Enter the playlist name on Spotify: ")
    user_id = sp.me()['id']
    spotify_playlist = sp.user_playlist_create(user_id, playlist_name, public=True)

    # Get the playlist ID
    playlist_id = spotify_playlist['id']

    # List to store not found songs
    not_found_songs = []

    # Iterate over each song and artist
    for song, artist in zip(songs, artists):
        try:
            # Search for the track on Spotify
            results = sp.search(q=f'{song} {artist}', type='track', limit=1)
            if results['tracks']['items']:
                # Add the first search result to the playlist
                track_uri = results['tracks']['items'][0]['uri']
                sp.user_playlist_add_tracks(user_id, playlist_id, [track_uri])
                print(f"Added '{song}' by '{artist}' to the playlist.")
            else:
                print(f"Track '{song}' by '{artist}' not found on Spotify.")
                not_found_songs.append((song, artist))
        except Exception as e:
            print(f"Error adding '{song}' by '{artist}' to the playlist:", e)

    print("Playlist creation complete!")

    # Save not found songs to another Excel file
    if not_found_songs:
        not_found_df = pd.DataFrame(not_found_songs, columns=['Song', 'Artist'])
        not_found_file_path = "not_found_songs.xlsx"
        not_found_df.to_excel(not_found_file_path, index=False)
        print(f"Not found songs saved to '{not_found_file_path}'.")

# Main program
if __name__ == "__main__":
    # Path to the Excel file
    file_path = "/FILE/PATH/HERE/FILE.XLSX"

    # Spotify access token
    access_token = "YOUR TOKEN HERE"

    # Create Spotify playlist from Excel file
    create_spotify_playlist_from_excel(file_path, access_token)
