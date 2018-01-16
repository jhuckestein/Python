from shutil import copyfile

template = 'Copy of TEMPLATE_CoursePlan_COURSEID_COURSETITLE_ndayV2.0.xlsm'
destination = 'CourseNamePlan.xlsm'
copyfile(template, destination)
print("Done")