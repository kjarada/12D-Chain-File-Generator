import pandas
import chianTxtIntro, ChainBodyStg2XML, ChainBodyStg3XML

pandas.set_option("display.max_rows", None)

pathToExcelFile = r'..\export.xlsx'
FSTG2 = open('PPW_STG2_Create_Polygons.chain',"w")
FSTG3 = open('PPW_STG3_Create_Polygons.chain',"w")


def read_execl_file_and_return_data(pathToExcelFile):

    wb = pandas.read_excel(pathToExcelFile)


    data = wb.loc[: , ["Lot No","Chainage Start (km)", "Chainage End (km)", "Activity", "Region", "Description", "Status", "Sub Area"]]
    #print(data.head())
    indexNames = data[ data['Chainage End (km)'] == 0 ].index

    data = data.drop(indexNames, inplace=False)

    indexNames = data[ data['Chainage End (km)'] > 1000 ].index
    data = data.drop(indexNames, inplace=False)

    indexNames = data[ data["Chainage Start (km)"] > 1000 ].index
    data = data.drop(indexNames, inplace=False)


    data = data.drop_duplicates(subset=['Chainage End (km)',"Chainage End (km)" ], keep=False)

    # for index, row in data.iterrows():
    #     print(row)
    return data
def iter_over_data_and_write_chain(data):
    recordNotStoredCounter = 0
    for index, row in data.iterrows():
        #STAGE 2
        if (str(row[4]).startswith("Stage 2")) & (str(row[6]) != "Abandoned") & ("Main Alignment" in str(row[7])):

            text = ChainBodyStg2XML.ChaintextStg2.replace("name1", (str(row[0]).replace(r"/", "-"))+" "+(row[5])+" "+ row[6])
            text = text.replace("ch1", str(row[1]*1000))
            text = text.replace("ch2", str(row[2]*1000))
            #Finding E3  
            if ("E3") in str(row[5]):

                if ("SF") in str(row[5]):

                    if ("Layer 1" or "layer 1") in str(row[5]):
                        text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "E3 SF " + "Layer 1")
                    elif ("Layer 2" or "layer 2") in str(row[5]):
                        text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "E3 SF " + "Layer 2")
                    else:
                        text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "E3 SF ")
                else:
                    if ("Layer 1" or "layer 1") in str(row[5]):
                        text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "E3 " + "Layer 1")
                    elif ("Layer 2" or "layer 2") in str(row[5]):
                        text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "E3 " + "Layer 2")
                    else:
                        text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "E3 ")

            #Finding C3          
            elif ("C3") in str(row[5]):
                if ("Layer 1" or "layer 1") in str(row[5]):
                    text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "C3 " + "Layer 1")
                elif ("Layer 2" or "layer 2") in str(row[5]):
                    text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "C3 " + "Layer 2")
                else:
                    text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "C3 ")
            else:
                if ("Layer 1" or "layer 1") in str(row[5]):
                    text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "Layer 1")
                elif ("Layer 2" or "layer 2") in str(row[5]):
                    text = text.replace("model1", "STG2 PPW BDY  Earthworks Filling "+ "Layer 2")
                else:
                    text = text.replace("model1", "STG2 PPW BDY " + str(row[3]))

            #Coloring Polygons
            if str(row[6]) == "Open":
                text = text.replace("colorr", "yellow")
            elif str(row[6]) == "Planned":
                text = text.replace("colorr", "blue")
            elif str(row[6]) == "Closed":
                text = text.replace("colorr", "green")
            else:
                text = text.replace("colorr", "red")

            FSTG2.write(text)
        #STAGE 3
        elif (str(row[4]).startswith("Stage 3")) & (str(row[6]) != "Abandoned") & ("Main Alignment" in str(row[7])):
            
            text = ChainBodyStg3XML.ChaintextStg3.replace("name1", (str(row[0]).replace(r"/", "-"))+" "+(row[5])+" "+ row[6])
            text = text.replace("ch1", str(row[1]*1000))
            text = text.replace("ch2", str(row[2]*1000))

            #Finding E3
            if ("E3") in str(row[5]):

                if ("SF") in str(row[5]):

                    if ("Layer 1" or "layer 1") in str(row[5]):
                        text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "E3 SF " + "Layer 1")
                    elif ("Layer 2" or "layer 2") in str(row[5]):
                        text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "E3 SF " + "Layer 2")
                    else:
                        text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "E3 SF")
                
                else:
                    if ("Layer 1" or "layer 1") in str(row[5]):
                        text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "E3 " + "Layer 1")
                    elif ("Layer 2" or "layer 2") in str(row[5]):
                        text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "E3 " + "Layer 2")
                    else:
                        text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "E3")
            #Finding C3
            elif ("C3") in str(row[5]):
                if ("Layer 1" or "layer 1") in str(row[5]):
                    text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "C3 " + "Layer 1")
                elif ("Layer 2" or "layer 2") in str(row[5]):
                    text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "C3 " + "Layer 2")
                else:
                    text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "C3 ")
            else:
                if ("Layer 1" or "layer 1") in str(row[5]):
                    text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "Layer 1")
                elif ("Layer 2" or "layer 2") in str(row[5]):
                    text = text.replace("model1", "STG3 PPW BDY Earthworks Filling "+ "Layer 2")
                else:
                    text = text.replace("model1", "STG3 PPW BDY " + str(row[3]))
            #Coloring Polygons
            if str(row[6]) == "Open":
                text = text.replace("colorr", "yellow")
            elif str(row[6]) == "Planned":
                text = text.replace("colorr", "blue")
            elif str(row[6]) == "Closed":
                text = text.replace("colorr", "green")
            else:
                text = text.replace("colorr", "red")
            FSTG3.write(text)
        else:
            print(row)
            recordNotStoredCounter+=1
    print(recordNotStoredCounter)

def main():

    FSTG2.write(chianTxtIntro.Intro)
    FSTG3.write(chianTxtIntro.Intro)


    iter_over_data_and_write_chain(data=read_execl_file_and_return_data(pathToExcelFile=pathToExcelFile)  )

    FSTG2.write(r'''</Commands>
    </Chain> 
    </xml12d>''')
    FSTG2.write(r'''</Commands>
    </Chain> 
    </xml12d>''')
    FSTG2.close
    FSTG3.close
    print("chain written")

main()