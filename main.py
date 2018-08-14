import json
import openpyxl

wb=openpyxl.load_workbook('''D:\RABOTA\easymark\easymark\All Work Items Easy-Mark Plus..xlsx''')
# sheet=wb.get_sheet_by_name('Test Plan')
sheet=wb.get_active_sheet()
epic=[]

path_to_folder="D:\RABOTA\easymark\easymark\easymark_workitems\\workitem."

epic_ids=[ path_to_folder+str(sheet.cell(row=i, column=2).value)+".json"
           for i in range(2,201) if sheet.cell(row=i, column=1).value=='Epic']

print(epic_ids)
data=''
for path in epic_ids[:1]:
    with open(path) as file:
        print("Epic ="+path)
        data = json.loads(file.read())
    with open("test.txt","w") as file:
        file.write("\tEPIC\n"+"\tuser_creator = "+data['dc:creator']['rdf:resource']+'\n')
        file.write("\n\t\tSTORY ->")
        story_id=[str(i['rdf:resource'].split('/')[-1]) for i in data['rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children']]
        for i in story_id:
            file.write("\n\t\t\t -"+i)
            print("Story=="+path_to_folder+str(i)+".json")
            with open(path_to_folder+str(i)+".json") as task_file:
                data_task = json.loads(task_file.read())
                dec = {str(i['rdf:resource'].split('/')[-1]) : i["oslc_cm:label"] for i in
                           data_task['rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children']}
                file.write("\n\t\t\t\t TASK ->" )
                for key,value in dec.items():
                    file.write("\n\t\t\t\t\t id ->"+key)
                    file.write("\n\t\t\t\t\t value ->"+ value)
                # print(data_task)

