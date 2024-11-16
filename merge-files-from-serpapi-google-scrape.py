import os
import pandas as pd

# 1. Get the current working directory (this is where the script is located)
current_dir = os.path.dirname(os.path.abspath(__file__))

# 2. List all Excel files in the current directory
file_list = [f for f in os.listdir(current_dir) if f.endswith('.xlsx')]

# 3. Initialize empty DataFrames for each category: Organic Results, Related Questions, and Related Searches
organic_results_df = pd.DataFrame()
related_questions_df = pd.DataFrame()
related_searches_df = pd.DataFrame()

# 4. Loop through each file in the file list and read the specific sheets
for file in file_list:
    file_path = os.path.join(current_dir, file)
    
    # Extract the seed keyword (part of the filename after 'organic_results_')
    keyword_query = file.split('organic_results_')[1].split('.')[0]

    try:
        # Read specific sheets from the current file
        organic_results = pd.read_excel(file_path, sheet_name='Organic Results')
        related_questions = pd.read_excel(file_path, sheet_name='Related Questions')
        related_searches = pd.read_excel(file_path, sheet_name='Related Searches')

        # Add the 'Keyword/Query' column to each dataframe
        organic_results['Keyword/Query'] = keyword_query
        related_questions['Keyword/Query'] = keyword_query
        related_searches['Keyword/Query'] = keyword_query

        # Append the data to the corresponding final DataFrame
        organic_results_df = pd.concat([organic_results_df, organic_results], ignore_index=True)
        related_questions_df = pd.concat([related_questions_df, related_questions], ignore_index=True)
        related_searches_df = pd.concat([related_searches_df, related_searches], ignore_index=True)

    except Exception as e:
        print(f"Error processing file {file}: {e}")

# 5. Save the combined data into separate Excel files
organic_results_df.to_excel(os.path.join(current_dir, 'Organic Results.xlsx'), index=False)
related_questions_df.to_excel(os.path.join(current_dir, 'Related Questions.xlsx'), index=False)
related_searches_df.to_excel(os.path.join(current_dir, 'Related Searches.xlsx'), index=False)

print("Merging complete! Files saved.")
