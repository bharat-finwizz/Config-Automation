import asyncio
from asyncio import sleep
from datetime import datetime
import re
from playwright.async_api import Playwright, async_playwright
import json

import pandas as pd

mydict = {}
mylist = []

df = pd.read_excel('test.xlsx', sheet_name='fromtenure')

        # Iterate over the columns and their rows that start with "other_details"
for column in df.columns:
    if column.startswith('other_details*') or column.startswith('applicant1*other_details*'):
        if not df[column].isnull().all():
            for index, value in df[column].iteritems():
                if not pd.isna(value):
                    if column.startswith('applicant_1*'):
                        print(f"Column {column} starts with 'applicant_1*'")
                    print(f'Value at index {index} in column {column}: {value}')
                    new_column_name = column
                    new_column_name = new_column_name.replace('*', '_')
                    new_column_name = re.sub(r'(_num|_str)', '', new_column_name)
                    new_column_name = re.sub(r'other_details_','', new_column_name)
                    print(new_column_name)
                    if df[column].dtype == 'float64':
                        value = int(value)
                    if new_column_name in mydict:
                        mydict[new_column_name].append(value)
                    else:
                        mydict[new_column_name] = [value]

                                                

with open('my_data.json', 'w') as file:
    json.dump(mydict, file) 
    file.close()
                                         

async def run(playwright: Playwright) -> None:
        browser = await playwright.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()


        await page.goto('https://fwzcallcenter1b.inforvio.in/index.php');
        await page.locator('[placeholder="Email-ID"]').nth(0).fill("pandiyan@propwiser.com")
        await page.locator('[placeholder="Password"]').fill("Loan1234!@")
        await page.locator('[class="btn btn-rose btn-round loginbtn"]').click()
        await sleep(5)
        await page.locator('text=Additional Fields Mgmt ').click()
        await sleep(1)
        
        count =1

        for key, values in mydict.items():
                print("Key:", key)
                for value in values:
                        print("Value:", value)
                        await page.locator('text=Code Master').click()
                        await page.locator('[class="btn btn-primary pull-right add_user"]').click()
                        await page.locator('[placeholder="Group Code"]').fill(key)
                        desc = key.replace('_', ' ').title()
                        
                        if isinstance(value, int):
                              break      
                        else: 
                                await page.locator('[placeholder="Key"]').fill(value)
                                await page.locator('[placeholder="Value"]').fill(value)              
                                # desc = desc.capitalize()
                                await page.locator('[placeholder="Description"]').fill(desc + ' = ' + value)
                        await page.locator('[id="foptnactive"]').select_option(label="Yes")
                        await page.locator('[class="btn btn-primary add_codemaster_btn"]').click()
                        await sleep(2)
                        
                await page.locator('text=Field Master').click()
                await page.locator('[class="btn btn-primary pull-right add_user"]').click()
                await page.locator('[placeholder="Code"]').fill(f'dd_{key}')
                if isinstance(value,int):
                        last_Value = mydict[key][-1]
                        text = ' less than ' 
                        await page.locator('[placeholder="Name"]').fill(f'{desc}({text}{last_Value}) ')
                        await page.locator('[id="fielddataType"]').select_option(label="integer")
                        if key.startswith('applicant1_'):
                                await page.locator('[id="fieldlevel"]').select_option(label="Applicant")
                        else:
                                        await page.locator('[id="fieldlevel"]').select_option(label="Aplication")
                else:
                        await page.locator('[placeholder="Name"]').fill(desc)
                        await page.locator('[id="fielddataType"]').select_option(label="dropdown")
                        if key.startswith('applicant1_'):
                                await page.locator('[id="fieldlevel"]').select_option(label="Applicant")
                        else:
                                await page.locator('[id="fieldlevel"]').select_option(label="Aplication")
                                await page.locator('[id="fieldgroupcode"]').select_option(key)
                await page.locator('[id="fieldismandatory1"]').click()
                await page.locator('[id="fieldactive"]').select_option(label="Yes")
                await page.locator('[class="btn btn-primary add_fields_btn"]').click()
                await sleep(2)
                await page.locator('text=Lender Mapping').click()
                await page.locator('[class="btn btn-primary pull-right add_user"]').click()
                await page.locator('[id="loan_category"]').select_option(label="Unsecured Loan")
                try:
                        await page.locator('[class="selectize-input items full has-options has-items"]').select_option(label="Aditya Birla Finance Ltd")
                except:
                        pass
                await page.locator('[id="fieldmapid"]').select_option(label=desc)
                await page.locator('[id="lenmapsequno"]').fill(str(count))
                count = count + 1
                await page.locator('[id="lenmapactive"]').select_option(label="Yes")
                await page.locator('[class="btn btn-primary add_lendermap_btn"]').click()
                await sleep(2)
                        
                        
                        
                                
                        
                        
                        
                        
                        
                        
                        
                        

        
        
        
        
        
        
        
        
        
    
        await context.close()
        await browser.close()

async def main() -> None:
    async with async_playwright() as playwright:
        await run(playwright)


asyncio.run(main())


