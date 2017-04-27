import os
def set_color(source_tab, dest_tab,r,g,b):
    with open('mb.txt','w') as fout:
        fout.write('set_color\n%s\n%s\n%d\n%d\n%d' % (source_tab, dest_tab,r,g,b))
    os.system('mb.mbx')