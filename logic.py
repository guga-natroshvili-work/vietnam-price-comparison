import pandas as pd
import xlsxwriter

def logic(ro_df,samo_df,date:str,pax:int, path:str):
    samo_df.rename(columns={'hotel': 'SamoHotel'}, inplace=True)
    samo_df.rename(columns={'adult': 'Adults'}, inplace=True)
    print(f'\n\n\n{path}\n\n\n')
    df = pd.merge(ro_df, samo_df, left_on=['HotelSamo','RoomCatName' ,'MealPlan', 'Adults'], right_on=['SamoHotel','Roomname', 'Meal', 'Adults'], how='left')
    
    ddf = df.drop(['HotelId', 'HotelSamo',
       'SamoHotel', 'hotelname', 'Roomname', 'Meal' ],axis=1)
    # add a column for the difference between the two prices
    ddf['PriceDiff'] = ddf['minPrice'] - ddf['Price']

    # rename columns
    ddf.rename(columns={'HotelName': 'Hotel', 'PlacementTypeName': 'Type','MealPlan':'Meal','RoomCatName': 'Room Category'}, inplace=True)
    save_path = path+f'/VPC {date},{pax}.xlsx'
    print(save_path)
    with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
        ddf.to_excel(writer, sheet_name='Sheet1', index=False)

        # Access the XlsxWriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Define a conditional format: highlight cells with Score > 90
        threshold = 0.000001  # Set a threshold for "near-zero" values

     
        worksheet.freeze_panes(1, 0)
        
        worksheet.conditional_format('F2:F10000', {
        'type': 'formula',
        'criteria': '=if($G2>0,$F2>$G2,"")',  # Formula: if G > max(H, I, J, K) for the same row
        'format': workbook.add_format({'bg_color': 'red', 'bold': True, 'font_color': 'white'})
        })
        worksheet.conditional_format('F2:F10000', {
        'type': 'formula',
        'criteria': '=$F2<$G2',  # Formula: if G > max(H, I, J, K) for the same row
        'format': workbook.add_format({'bg_color': 'green', 'bold': True, 'font_color': 'white'})
        })
        worksheet.conditional_format('L2:L10000', {
        'type': 'formula',
        'criteria': '=if($G2>0,$F2>$G2,"")',  # Formula: if G > max(H, I, J, K) for the same row
        'format': workbook.add_format({'bg_color': 'red', 'bold': True, 'font_color': 'white'})
        })
        worksheet.conditional_format('L2:L10000', {
        'type': 'formula',
        'criteria': '=$F2<$G2',  # Formula: if G > max(H, I, J, K) for the same row
        'format': workbook.add_format({'bg_color': 'green', 'bold': True, 'font_color': 'white'})
        })
        for col_idx, col in enumerate(ddf.columns):
            col_length = len(col)
            max_length = len(ddf[col].astype(str).loc[ddf[col].astype(str).str.len().idxmax()])
            if col_length > max_length:
                worksheet.set_column(col_idx, col_idx, col_length + 2)  # Adjust width with padding

            else:
            
                worksheet.set_column(col_idx, col_idx, max_length + 2)  # Adjust width with padding


        worksheet.set_column('G:G', None, None, {'hidden': True})
        
