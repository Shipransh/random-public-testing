import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from collections import defaultdict

# Your clustering function
def text_clustering_df_definedLabels(df, number_of_clusters, random_key_state):
    vec = TfidfVectorizer(stop_words='english', ngram_range=(1,3))
    vec.fit(df['Solution'].values)
    features = vec.transform(df['Solution'].values)
    
    clust = KMeans(init='k-means++', n_clusters=number_of_clusters, n_init=10, random_state=random_key_state)
    clust.fit(features)
    
    yhat = clust.predict(features)
    df['Cluster_Labels'] = clust.labels_
    
    # Include Engineer Name, Part Number, and Cluster Number in Cluster Labels
    df['Cluster_Labels'] = df['Engineer Name'] + "_" + df['Part_Number'] + "_Cluster_" + df['Cluster_Labels'].astype(str)
    
    df_engineer_w_labels = df.sort_values('Cluster_Labels').reset_index(drop=True)
    return df_engineer_w_labels

# Main script to apply the clustering function for each engineer and part number combination
def apply_clustering_to_all(df, random_key_state=42):
    # Group by Engineer and Part_Number and count the number of cases
    engineer_part_counts = df.groupby(['Engineer Name', 'Part_Number']).size().reset_index(name='Case_Count')
    
    # Sort by Case_Count, Engineer Name, and Part Number
    engineer_part_counts = engineer_part_counts.sort_values(by=['Case_Count', 'Engineer Name', 'Part_Number'], ascending=[False, True, True])
    
    # Initialize a dictionary to hold clusters for each Engineer-Part_Number combination
    all_clusters = pd.DataFrame()  # To store all clusters

    # Loop through each Engineer and their associated Part Numbers
    for engineer in engineer_part_counts['Engineer Name'].unique():
        engineer_cases = df[df['Engineer Name'] == engineer]
        
        # Loop through each Part Number for the engineer
        for _, row in engineer_part_counts[engineer_part_counts['Engineer Name'] == engineer].iterrows():
            part_number = row['Part_Number']
            part_cases = engineer_cases[engineer_cases['Part_Number'] == part_number]
            
            # Determine the number of clusters (you may adjust this logic)
            number_of_clusters = min(len(part_cases), 5)  # Example: cap clusters at 5 or less if fewer cases
            
            # Apply the clustering function
            clustered_df = text_clustering_df_definedLabels(part_cases, number_of_clusters, random_key_state)
            
            # Append the results to the all_clusters DataFrame
            all_clusters = pd.concat([all_clusters, clustered_df], ignore_index=True)
    
    return all_clusters

# Example usage
# Assuming df is your DataFrame with the relevant columns
# final_clusters = apply_clustering_to_all(df)
