import sys,os,subprocess,datetime
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point, Polygon, MultiPolygon, LineString, MultiLineString
from IPython.display import HTML
import matplotlib.pyplot as plt
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor
    

WKT_SWISS="""PROJCS["unnamed",
GEOGCS["unnamed",
    DATUM["Switzerland_CH_1903",
        SPHEROID["Bessel 1841",6377397.155,299.1528128],
        TOWGS84[660.077,13.551,369.344,-0.804816,-0.577692,-0.952236,5.66]],
    PRIMEM["Greenwich",0],
    UNIT["degree",0.0174532925199433]],
PROJECTION["Swiss_Oblique_Cylindrical"],
PARAMETER["latitude_of_center",46.9524055555],
PARAMETER["central_meridian",7.4395833333],
PARAMETER["false_easting",600000],
PARAMETER["false_northing",200000],
UNIT["Meter",1.0]]
"""

WKT_WGS="""GEOGCS["unnamed",
    DATUM["WGS_1984",
        SPHEROID["WGS 84",6378137,298.257223563],
        TOWGS84[0,0,0,0,0,0,0]],
    PRIMEM["Greenwich",0],
    UNIT["degree",0.0174532925199433]]
"""

MIF_SWISS='CoordSys Earth Projection 25, 1003, "m", 7.4395833333, 46.9524055555, 600000, 200000'
MIF_WGS='CoordSys Earth Projection 1, 104'


def run_nb(ju_nb):
    """Execute a jupyter notebook"""

    nb = nbformat.read(open(ju_nb), as_version=4)
    ep = ExecutePreprocessor(timeout=600, kernel_name='python3')
    ep.preprocess(nb, {'metadata': {'path': os.path.dirname(ju_nb)}})
    nbformat.write(nb, open(ju_nb, mode='wt'))
    
def main():
    if len(sys.argv)==2:
        nb=sys.argv[1]
        print('starting',nb)
        run_nb(nb)
    else:
        print('usage: mio.py notebook.jpynb to execute as jupyter notebook')

if __name__ == "__main__":
    main()

# set automatic oprions on import
def set_options():
    from IPython.core.interactiveshell import InteractiveShell
    InteractiveShell.ast_node_interactivity = "all"
    pd.set_option("display.max_rows",100)
    pd.set_option("display.max_columns",100)
    pd.set_option("display.max_colwidth",100)

    plt.style.use('ggplot')
    plt.style.use('seaborn-colorblind')
    
set_options()


def geom_point(xy):
    return Point(xy)

def geom_square(xy_center,d):
    r=d/2
    x,y=xy_center
    p0=(x-r,y-r)
    p1=(x-r,y+r)
    p2=(x+r,y+r)
    p3=(x+r,y-r)
    return Polygon([p0,p1,p2,p3])

# create normed address key for matching
def adr_key(zi,street,no):
    zi=str(zi)
    street=str(street)
    no=str(no)
    
    # norm street
    if ',' in  street:
        l= street.split(',')
        street = ''.join(reversed(l))
        
    
    # add space for simpler filtering
    street=street.lower().strip()
    street=' %s ' % street
    
    # replace abbr.
    street=street.replace('str.','strasse')
    street=street.replace('ch.','chemin')
    street=street.replace('rte.','route')
    street=street.replace('bvd.','boulevard')
    street=street.replace('sent.','sentier')
    street=street.replace('av.','avenue')
    street=street.replace('pl.','place')
    street=street.replace('imp.','impasse')
    # if point is  forgotten
    street=street.replace(' str ','strasse')
    street=street.replace(' ch ','chemin')
    street=street.replace(' rte ','route')
    street=street.replace(' bvd ','boulevard')
    street=street.replace(' sent ','sentier')
    street=street.replace(' av ','avenue')
    street=street.replace(' pl ','place')
    street=street.replace(' imp ','impasse')


    street=''.join([c if c. isalnum() else '' for c in street])

    # norm house number
    no=''.join([c if c. isdigit() else '' for c in no])
    s=''.join(no)
    if not s:
        s='0'
    no=str(int(s))

    return '%s_%s_%s' % (zi,street,no)

def flatten(df):
    """ Make simple dataframe without multi columns"""
    header=[]
    for t in df.columns.values:
        if type(t)==str:
            header.append(t)
        else:
            header.append('_'.join(t).strip('_'))
    df_ret=df.copy()
    df_ret.columns=header
    df_ret=df_ret.reset_index()
    return df_ret


def now():
    return datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')

def now2():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def today():
    return datetime.datetime.now().strftime('%Y-%m-%d')

def write_tab(gdf,tab_name,crs_wkt=WKT_SWISS):
    
    gdf=gdf.copy()
    
    # bring multi to reduce object types (Fiona can save only on)
    def to_multi(geom):
        if geom.type=='Polygon':
            return MultiPolygon([geom])
        elif geom.type=='Line':
            return MultiLine([geom])        
        else:
            return geom
    
    gdf.geometry=[to_multi(geom) for geom in gdf.geometry]
    
    # make the columns fit for Mapinfo
    new_cols=[]
    for s in gdf.columns:
        s=''.join([c if c.isalnum() else '_' for c in s])
        for i in range(5):
            s=s.replace('__','_')
        s=s.strip('_')
        s=s[0:30]

        new_cols.append(s)
    gdf.columns=new_cols
               
    # create my own schema (without schema all strings are 254)
    props={}
    for col,typ in gdf.dtypes.iteritems():
        if col!=gdf.geometry.name:
            if str(typ).startswith('int'):
                styp='int'
            elif str(typ).startswith('float'):
                styp='float'
            else:
                gdf[col]=gdf[col].astype('str')  
                max_len=gdf[col].map(len).max()
                styp='str:%d' % max_len
            props[col]=styp
            
    schema={}
    # set geometry type of the first object for the whole layer
    schema['geometry']= geo_obj_type=gdf.geometry.iloc[0].geom_type
    schema['properties']=props
       
    # delete files if already there, otherwise an error is raised
    base_dest,ext_dest= os.path.splitext(tab_name)
    if ext_dest.lower()=='.tab':
        ext_list=['.tab','.map,','.dat','.id']
    elif ext_dest.lower()=='.mif':
        ext_list=['.mif','.mid']
    else:
        sys.exit("ERROR: extension of '%s' should be .tab or .mif." % tab_name)
    
    for ext in ext_list:
        #print("removing %s" % base_dest+ext)
        delete_if_exists(base_dest+ext)

    gdf.to_file(tab_name,driver='MapInfo File',crs_wkt=crs_wkt,schema=schema)    
    return print(len(gdf), 'rows of type', geo_obj_type, 'written to mapinfo file.', now2())
    
def read_grid(sFile):
    # read header
    dic={}
    fin=open(sFile)
    iHeader=0
    for i in range(8):
        l=next(fin).strip().split()
        if len(l)>2:
            break
        iHeader=iHeader+1
        dic[l[0].lower()]=l[1]

    df_data=pd.read_csv(sFile,skiprows=iHeader,delim_whitespace=True,header=None)
    
    # set indexes and columns
    
    ncols=int(dic['ncols'])
    nrows=int(dic['nrows'])
    xllcorner=float(dic['xllcorner'])
    yllcorner=float(dic['yllcorner'])
    cellsize=float(dic['cellsize'])
    
    lCols=[xllcorner+i*cellsize for i in range(ncols)]
    ystart=yllcorner+nrows*cellsize
    df_data.columns=lCols
    lRows=[ystart-i*cellsize for i in range(nrows)]
    df_data.index=lRows
    
    return df_data


def write_grid(df,sFile,no_data=0):
    "write ESRI grid files"
    df=df.copy()
    nrows, ncols= df.shape
    x0=df.columns.tolist()[0]
    x1=df.columns.tolist()[-1]
    
    y0=(df.index.tolist()[-1])
    y1=(df.index.tolist()[0])
    
    cs=(x1-x0)/(ncols-1)
    
    fout=open(sFile,"w")
    fout.write("ncols %d\n" % ncols)
    fout.write("nrows %d\n" % nrows)
    
    fout.write("xllcorner %f\n" % x0)
    fout.write("yllcorner %f\n" % (y0-cs))
    fout.write("cellsize %f\n" % cs)
    fout.write("nodata_value %d\n" % no_data)

    df.to_csv(fout,sep=" ",header=None,index=None)
    return 'ascii grid written.'
    
def read_mif(sMif):
    sBase=os.path.splitext(sMif)[0]
    fmif=open(sBase+".mif",encoding='latin-1')
    
    # read Delimiter
    s=""
    while(not s.lower().startswith('delimiter')):
        s=next(fmif) 
        sDelimiter=s.split()[1].strip('"').strip("'")

    # read header
    s=""
    while(not s.lower().startswith('columns')):
        s=next(fmif)
    iCols=int(s.split()[1])
    lColNames=[]
    for i in range(iCols):
        s=next(fmif).strip()
        sColName=s.split()[0]
        lColNames.append(sColName)
    fmif.close()
    
    # create dataframe and read mid file
    df=pd.read_csv(sBase+".mid",sep=sDelimiter,header=None,encoding='latin-1')
    df.columns=lColNames
    
    # clean uo
    fmif.close()
    
    # give back dtaframe
    return df

def write_mif(df,sMif,x=0,y=0,sCoordSys='swiss'):
    """ Write mif, but only points implemented"""
    df=df.copy()
    sSep=";"
    sFileTitle=os.path.splitext(sMif)[0]
    
    dCoordSys={}
    dCoordSys['swiss']='CoordSys Earth Projection 25, 1003, "m", 7.4395833333, 46.9524055555, 600000, 200000'
    dCoordSys['wgs84']='CoordSys Earth Projection 1, 104'
    if sCoordSys in dCoordSys: sCoordSys=dCoordSys[sCoordSys]
        
    lColumns=[]
    dColNames={}
    for sFieldName in df:
        if not sFieldName.startswith("mi_"): # skip the mai_fields
            series=df[sFieldName]
            sType = str(series.dtype)
          
            # Make fieldnames fit to mapinfo
            sClean=""
            for c in sFieldName:
                if (ord(c) in range(ord('A'),ord('z')) or (ord(c) in range(ord('0'),ord('9')))):
                    sClean=sClean+c
                else:
                    sClean=sClean+"_"
            sFieldName=sClean[0:30]
                    
            i=0
            while(sFieldName in dColNames):
                s=str(i) 
                sFieldName=sFieldName[0:len(s)]+s
                i=i+1
            dColNames[sFieldName]=1
            if "int" in sType:
                lColumns.append("%s Integer" % sFieldName)
            elif "float" in sType:
                lColumns.append("%s Float" % sFieldName)
            else:
                iLen=int(series.astype(str).map(len).max())
                lColumns.append("%s Char(%d)" % (sFieldName,iLen))
     
    # write mif file header
    fmif=open(sFileTitle+".mif","w")
    fmif.write('Version 300\n')
    fmif.write('Charset "Neutral"\n')
    fmif.write('Delimiter "%s"\n' % sSep)
    fmif.write('%s\n' % sCoordSys)
    fmif.write('Columns %d\n' % len(lColumns))
    for sCol in lColumns:
        fmif.write("\t%s\n" % sCol)
    fmif.write('Data\n\n')
        
    # write objects into mif

    for index,row in df.iterrows():
        if type(x)==str:
            fx=float(row[x])
            fy=float(row[y])
        else:
            fx=x
            fy=y
        s="Point %f %f\n" % (fx,fy)
        fmif.write(s)
    fmif.close()

    # write mid file
    df.to_csv(sFileTitle+".mid",sep=sSep,header=None,index=None)
    return 'mif file written.'

def search_files(search_path):
    """Search all files and return a datafram"""
    dic={'dir':[],'file':[]}
    for root,dirs,files in os.walk(search_path):
        dic['dir'].append(root)
        dic['file'].append('.')
        for file in files:
            dic['file'].append(file)
            dic['dir'].append(root)
    return pd.DataFrame(dic)

def delete_if_exists(file_name):
    if os.path.isfile(file_name):
        os.remove(file_name)
        
def swiss_wgs(sX,sY):
    sX=str(sX)
    sY=str(sY)

    if sX.strip()[0].isdigit():
        x=float(sX)
        y=float(sY)
        x1 = (x - 600000) / 1000000
        y1 = (y - 200000) / 1000000
        L = 2.6779094 + 4.728982 * x1 + 0.791484 * x1 * y1 + 0.1306 * x1 * y1 * y1 - 0.0436 * x1 * x1 * x1
        p = 16.9023892 + 3.238272 * y1 - 0.270978 * x1 * x1 - 0.002528 * y1 * y1 - 0.0447 * x1 * x1 * y1 - 0.014 * y1 * y1 * y1
        return (L*100/36,p*100/36)
    else:
        return (-1,-1)
    
def wgs_swiss(sLon,sLat):
    sLon=str(sLon)
    sLat=str(sLat)

    if sLon.strip()[0].isdigit():
        Lon=float(sLon)
        Lat=float(sLat)
        p = (Lat * 3600 - 169028.66) / 10000
        L = (Lon * 3600 - 26782.5) / 10000
        y = 200147.07 + 308807.95 * p + 3745.25 * L * L + 76.63 * p * p - 194.56 * L * L * p + 119.79 * p * p * p
        x = 600072.37 + 211455.93 * L - 10938.51 * L * p - 0.36 * L * p * p - 44.54 * L * L * L
        return (x,y)
    else:
        return (-1,-1)
    
def run_mb(mb_script):
    """ run a Mapbasic Script, mapinfo and mapbasic folders have to be defined in the environment variable PATH"""
    delete_if_exists('mb.mb')
    delete_if_exists('mb.mbx')
    delete_if_exists('mb.err')
    with open('mb.mb','w') as fout: 
        fout.write(mb_script)
    subprocess.run('mapbasic.exe -D mb.mb',stdout=subprocess.PIPE,encoding='latin-1')
    if os.path.isfile('mb.err'):
        print('Error compiling mb.mb:')
        with open('mb.err') as ferr:
            for s in ferr:
                print(s)
    else:
        subprocess.run('mapinfow.exe mb.mbx')
    
