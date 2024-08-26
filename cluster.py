import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans

def vectorize_solutions(df, fill_value=""):
    """
    Vectorize the Solution column for the entire DataFrame,
    handling NaN values by filling them with a specified value.
    """
    # Fill NaN values in the Solution column with a specified fill_value (e.g., empty string)
    df['Solution'] = df['Solution'].fillna(fill_value)
    
    vec = TfidfVectorizer(stop_words='english', ngram_range=(1,3))
    features = vec.fit_transform(df['Solution'].values)
    return vec, features

def apply_clustering(df, features, number_of_clusters, random_key_state):
    """Apply KMeans clustering on the provided features."""
    clust = KMeans(init='k-means++', n_clusters=number_of_clusters, n_init=10, random_state=random_key_state)
    clust.fit(features)
    return clust.labels_

def text_clustering_with_labels(df, features, number_of_clusters, random_key_state):
    """Cluster the data and add cluster labels."""
    df = df.copy()
    df.loc[:, 'Cluster_Labels'] = apply_clustering(df, features, number_of_clusters, random_key_state)
    df.loc[:, 'Cluster_Labels'] = df['Engineer Name'] + "_" + df['Part_Number'] + "_Cluster_" + df['Cluster_Labels'].astype(str)
    df = df.sort_values('Cluster_Labels').reset_index(drop=True)
    return df

def apply_clustering_to_all(df, random_key_state=42):
    # Vectorize the entire DataFrame once, handling NaN values
    vec, features = vectorize_solutions(df)
    
    # Group by Engineer and Part_Number and count the number of cases
    engineer_part_counts = df.groupby(['Engineer Name', 'Part_Number']).size().reset_index(name='Case_Count')
    
    # Sort by Case_Count, Engineer Name, and Part Number
    engineer_part_counts = engineer_part_counts.sort_values(by=['Case_Count', 'Engineer Name', 'Part_Number'], ascending=[False, True, True])
    
    # Initialize a DataFrame to hold all the clustered results
    all_clusters = pd.DataFrame()
    
    # Loop through each Engineer and their associated Part Numbers
    for engineer in engineer_part_counts['Engineer Name'].unique():
        engineer_cases = df[df['Engineer Name'] == engineer]
        
        # Get the vectorized features for the engineer's cases
        engineer_features = features[df['Engineer Name'] == engineer]
        
        # Loop through each Part Number for the engineer
        for _, row in engineer_part_counts[en
