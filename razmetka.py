import pandas as pd
import os
from tqdm import tqdm

def process_dialogs():
    result_df = pd.DataFrame(columns=[
        'groupId',
        'taskMetadata.prop1',
        'taskData.dialog',
        'priority',
        'expirationTime'
    ])
    
    dialogs_dir = 'dialogs'
    
    for i in tqdm(range(1, 134), desc="Processing dialogs"):
        file_path = os.path.join(dialogs_dir, f"{i}.xlsx")
        
        try:
            dialog_df = pd.read_excel(file_path)
            
            dialog_text = []
            for _, row in dialog_df.iterrows():
                role = row['Role'].lower()
                content = row['Content']
                if pd.notna(content):
                    dialog_text.append(f"{role}: {content}")
            
            full_dialog = "\n".join(dialog_text)
            
            result_df = pd.concat([result_df, pd.DataFrame([{
                'groupId': None,
                'taskMetadata.prop1': None,
                'taskData.dialog': full_dialog,
                'priority': None,
                'expirationTime': None
            }])], ignore_index=True)
            
        except Exception as e:
            print(f"Error processing file {i}.xlsx: {str(e)}")

    result_df.to_excel("consolidated_dialogs.xlsx", index=False)
    print("Processing completed! Saved to consolidated_dialogs.xlsx")

if __name__ == "__main__":
    process_dialogs()