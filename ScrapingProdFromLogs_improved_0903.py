import os
import json
import pandas as pd
import numpy as np

def get_preludes_after_second_setup_phase(turn_logs):
    setup_phases = [log for log in turn_logs if log['phaseType'] == 'PlayerSetup']
    if len(setup_phases) < 2:
        return {}
    return {player_id: player_data['selectedPreludeCards'] for player_id, player_data in setup_phases[1]['playerInfo'].items()}

def get_gen2_data(turn_logs):
    gen2_data = next((log for log in turn_logs if log["generation"] == 2), {})
    player_info = gen2_data.get('playerInfo', {})
    return player_info

def extract_generation_data(turn_logs, generation_number, player_id):
    gen_data = next((log for log in turn_logs if log["generation"] == generation_number), None)
    if not gen_data:
        return {}
    player_info = gen_data.get('playerInfo', {}).get(str(player_id), {})
    return {
        f'Generation {generation_number} terraFormingRating': player_info.get('score', {}).get('terraFormingRating', None),
        f'Generation {generation_number} MC production': player_info.get('resourceData', {}).get('mc', {}).get('production', None),
        f'Generation {generation_number} Steel production': player_info.get('resourceData', {}).get('steel', {}).get('production', None),
        f'Generation {generation_number} Titanium production': player_info.get('resourceData', {}).get('ti', {}).get('production', None)
    }

# Directory containing the JSON files
directory = 'C:\\Program Files (x86)\\Steam\\steamapps\\common\\Terraforming Mars\\Logs\\GameLogs'

# Load the support data
support_data_corps = pd.read_excel('C:\\Users\\Andi\\Desktop\\data analytcs\\TM_Analysis\\Support data for Template .xlsx', sheet_name='Corporation starting resources')
support_data_preludes = pd.read_excel('C:\\Users\\Andi\\Desktop\\data analytcs\\TM_Analysis\\Support data for Template .xlsx', sheet_name='Preludes starting resources')

# Create a DataFrame to store the extracted data
df = pd.DataFrame()

# Loop through each file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.json'):
        
        # Load the JSON data from the file
        file_path = os.path.join(directory, filename)
        with open(file_path, 'r') as f:
            json_data = json.load(f)
        
        # Extract the game length (number of generations)
        game_length = json_data['turnLogs'][-1]['generation']

        # Extract the player names
        player_names = [player['name'] for player in json_data['players'].values()]
        
        # Extract the elo
        elo = [int(player['elo']) for player in json_data['players'].values()]
        
        # Extract the game id
        game_id = [str(json_data['id'])] * len(player_names)
        
        # Extract the game board type
        game_board = [json_data['boardType']] * len(player_names)
        
        # Extract the game start time
        game_start_time = [json_data['gameStartTime']] * len(player_names)
        
        # List version of game length for the dataframe
        game_length_list = [game_length] * len(player_names)        

        # Extract the corporations
        corporations = [player['corporation'] for player in json_data['players'].values()]
        
        # Extract the final scores
        final_plant_conversion_phase = [phase for phase in json_data['turnLogs'] if phase.get('phaseType') == 'FinalPlantConversion']
        final_scores = [player_info['score']['finalScore'] for player_info in final_plant_conversion_phase[0]['playerInfo'].values()]

        # Extracting placement in the game:
        placements = pd.Series(final_scores).rank(ascending=False, method='min').astype(int).tolist()

        # Get the starting MC, steel, and titanium for each player from the support data
        starting_mc = [support_data_corps[support_data_corps['Corporation'] == corp]['Starting MC'].values[0] for corp in corporations]
        starting_steel = [support_data_corps[support_data_corps['Corporation'] == corp]['Starting Steel resources'].values[0] for corp in corporations]
        starting_titanium = [support_data_corps[support_data_corps['Corporation'] == corp]['Starting Titanium resources'].values[0] for corp in corporations]
        
        # Extract the preludes
        prelude_data = get_preludes_after_second_setup_phase(json_data['turnLogs'])
        prelude_1 = [prelude_data.get(str(i+1), [None, None])[0] if len(prelude_data.get(str(i+1), [None])) > 0 else None for i in range(len(player_names))]
        prelude_2 = [prelude_data.get(str(i+1), [None, None])[1] if len(prelude_data.get(str(i+1), [None])) > 1 else None for i in range(len(player_names))]

        # Get the prelude resources
        prelude_1_resources = support_data_preludes.set_index('Prelude').reindex(prelude_1).reset_index().fillna('')
        prelude_2_resources = support_data_preludes.set_index('Prelude').reindex(prelude_2).reset_index().fillna('')

        # Extract Generation 2 data
        gen2_data = get_gen2_data(json_data['turnLogs'])
        gen2_tr = [gen2_data[str(i+1)]['score']['terraFormingRating'] for i in range(len(player_names))]
        gen2_mc_prod = [gen2_data[str(i+1)]['resourceData']['mc']['production'] for i in range(len(player_names))]
        gen2_steel_prod = [gen2_data[str(i+1)]['resourceData']['steel']['production'] for i in range(len(player_names))]
        gen2_ti_prod = [gen2_data[str(i+1)]['resourceData']['ti']['production'] for i in range(len(player_names))]

        # Append the extracted data to the DataFrame
        for player_index, player_name in enumerate(player_names):
            data_for_df = {
                'Player Name': player_name,
                'Elo': elo[player_index],
                'Game ID': game_id[player_index],
                'Board': game_board[player_index],
                'Game Start Time': game_start_time[player_index],
                'Game Length (Generations)': game_length_list[player_index],
                'Final Score': final_scores[player_index],
                'Placement': placements[player_index],
                'Corporation': corporations[player_index],
                'Starting MC': starting_mc[player_index],
                'Starting Steel': starting_steel[player_index],
                'Starting Titanium': starting_titanium[player_index],
                'Prelude 1': prelude_1[player_index],
                'Prelude 2': prelude_2[player_index],
                'Prelude 1 MC': prelude_1_resources['MC'][player_index],
                'Prelude 1 Steel': prelude_1_resources['Steel'][player_index],
                'Prelude 1 Titanium': prelude_1_resources['Titanium'][player_index],
                'Prelude 2 MC': prelude_2_resources['MC'][player_index],
                'Prelude 2 Steel': prelude_2_resources['Steel'][player_index],
                'Prelude 2 Titanium': prelude_2_resources['Titanium'][player_index],
                'Generation 2 terraFormingRating': gen2_tr[player_index],
                'Generation 2 MC production': gen2_mc_prod[player_index],
                'Generation 2 Steel production': gen2_steel_prod[player_index],
                'Generation 2 Titanium production': gen2_ti_prod[player_index]
            }

            for gen in range(2, game_length + 1):  # Start from Generation 2
                gen_info = extract_generation_data(json_data['turnLogs'], gen, player_index+1)
                data_for_df.update(gen_info)

            temp_df = pd.DataFrame([data_for_df])
            df = pd.concat([df, temp_df])

        # Reorder columns just before saving to Excel
        columns_order = [
            "Player Name", "Elo", "Game ID", "Board", "Game Start Time", 
            "Game Length (Generations)", "Final Score", "Placement", "Total Production", "MC/VP", 
            "Corporation", "Starting MC", "Starting Steel", "Starting Titanium",
            "Prelude 1", "Prelude 2", "Prelude 1 MC", "Prelude 1 Steel", "Prelude 1 Titanium",
            "Prelude 2 MC", "Prelude 2 Steel", "Prelude 2 Titanium"
            ]

        # For each generation, you need to add terraFormingRating, MC production, Steel production, and Titanium production.
        for i in range(2, 15): # From generation 2 to 14
            columns_order.append(f"Generation {i} terraFormingRating")
            columns_order.append(f"Generation {i} MC production")
            columns_order.append(f"Generation {i} Steel production")
            columns_order.append(f"Generation {i} Titanium production")

        # Ensure all columns in the desired order exist in the DataFrame
        for col in columns_order:
            if col not in df.columns:
                df[col] = np.nan

        # Reorder the DataFrame columns
        df = df[columns_order]

# Save the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
