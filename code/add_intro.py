import re


def extract_identifier(text,company_name):
    match = re.search(r'\((.*?)\)', text)
    if match:
        return f"{match.group(1)} by {company_name}", match.group(1)
    return text, ""



def intro_information(slide,row, date_en, date_cn,date_idx,description_idx,product_idx,category_idx,language):

    company_name = row['Brand or Association']
    #product text
    product_placeholder = slide.placeholders[product_idx]
    product_text, image_name = extract_identifier(row['RESET ID'],company_name)
    product_placeholder.text = product_text
    recyclable_content = int(format(float(row['RC'].rstrip("%")),'.0f'))
    if language == "EN":

        date_text = f"As of {date_en}"
        description = f"{image_name} is a carbon neutral carpet tile that is made from {recyclable_content}% recycled content, is 100% recyclable and contains no harmful red-listed chemicals."
        
    elif language == "CN":
        
        date_text = f"{date_cn}"
        description = f"{image_name} 是一种碳中和方块地毯，由{recyclable_content}%的可回收成分制成且100%可回收，不含有害的红色清单化学成分。"      

    #date
    date_placeholder = slide.placeholders[date_idx]
    date_placeholder.text = date_text

    #description
    description_placeholder = slide.placeholders[description_idx]
    description_placeholder.text = description



    #category
    category_text = row['Series']
    category_placeholder = slide.placeholders[category_idx]
    category_placeholder.text = category_text   

