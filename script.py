import clr
from pprint import pprint
from airtable import Airtable
__fullframeengine__ = True
import csv
from System.Collections.Generic import List

import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
import clr


#get the base key
base_key = "appzkwGZTCmkWKlfJ"

#get the table name
table_name = "Project_Base"

#get the api key from the Airtable api documentation
api_key = "keyQWE2H4jycRNW1x"

#initiate the Airtable Base - Setting the Airtable base
airtable = Airtable(base_key,table_name, api_key)

#Converting the individual lines from Schedule dataframe to insertable records - single line records

#talking to revit

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

model_pvd = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME))
rulev = DB.FilterStringBeginsWith()
sku_prefix = 'WWI'
model_rule = DB.FilterStringRule(model_pvd, rulev, sku_prefix, False)
model_filter = DB.ElementParameterFilter(model_rule)


#cl = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()

fl = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(model_filter).ToElements()





fl_id = []
for item in fl:
	fl_id.append(item.Id)

fl_list= []
for item in fl:
	fl_list.append(item.Name)

SKU_Set = set(fl_list)
schedule_dict = []

project_name = doc.ProjectInformation.Name.ToString() #DB.ElementId(DB.BuiltInParameter.PROJECT_NAME)

for item in SKU_Set:
	schedule_dict.append({'WWI-SKU':item,'Quantity':fl_list.count(item), 'Project': project_name })


#intitate Dlete record placeholder if it exist -- This will contain the records ID to delete
to_delete = []

#initiate delete Project placeholder if exist
to_delete_project = schedule_dict[1]['Project']

for id in airtable.get_all():
   if id['fields']['Project'] == to_delete_project:
                  to_delete.append((id['id']))


# Batch delete for the existing record from the same project
airtable.batch_delete(to_delete)


#Inserting all the records from schedule to 'project base'
#for item in schedule_dict:
for item in schedule_dict:
   airtable.insert(item,typecast=True)
   
UI.TaskDialog.Show("Master Schedule Upload", "YOUR PROJECT IS SUCCESSFULLY UPLOADED ON THE AIRTABLE")