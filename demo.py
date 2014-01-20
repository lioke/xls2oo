#NB only for test
import os
os.chdir("D:\\workspace\\xls2oo")

from xls2oo import *

xlsobj = XLS2OO("demo.xls")
#t_anagrafica = xlsobj.objects.filter(table_name="demo_anagrafica")
#t_anagrafica = t_anagrafica.objects[0]

try:
    t_anagrafica = xlsobj.get_table("demo_anagrafica")
except:
    pass

filtered = t_anagrafica.objects.filter(id=1)
filtered = t_anagrafica.objects.filter(id__ge=1)

#filter arguments can be filed__ + one of these:
    #exact
    #iexact
    #contains
    #icontains
    #startswith
    #istartswith
    #endswith
    #iendswith
    #lt
    #le
    #gt
    #ge
    #in
first = filtered[0]
for f in filtered:
    print f.pk
    print f.cap
    print f.pk.value
    print f.pk.type
    #print f.pk.type

try:
    t_anagrafica = xlsobj.get(table_name="demo_anagrafica")
except:
    pass
pass
