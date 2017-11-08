USE_GEOPANDAS = True

import sys,os,datetime
import pandas as pd
import numpy as np

if USE_GEOPANDAS:
    import geopandas as gpd
    from shapely.geometry import Point, Polygon, MultiPolygon, LineString, MultiLineString, shape
    import rasterio
    import rasterio.features
    
from IPython.display import HTML
import matplotlib.pyplot as plt
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor
import win32com.client

import affine


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
    if len(sys.argv)>2:
        os.environ["JUPYTER_PARAMETER"] = sys.argv[2]
    else:
        os.environ["JUPYTER_PARAMETER"] = ""
    # the jupyter notebook can retrieve os.environ["JUPYTER_PARAMETER"]
    nb = nbformat.read(open(ju_nb), as_version=4)
    ep = ExecutePreprocessor(timeout=600, kernel_name='python3')
    ep.preprocess(nb, {'metadata': {'path': os.path.dirname(ju_nb)}})
    # write sometimes destroys the notebook!
    # better just write a copy
    new_nb_name=os.path.splitext(ju_nb)[0]+'_last_run.ipynb'
    nbformat.write(nb, open(new_nb_name, mode='wt'))
    
def main():
    if len(sys.argv)>1:
        nb = sys.argv[1]
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
    pd.set_option("display.max_colwidth",1024)

    plt.style.use('ggplot')
    plt.style.use('seaborn-colorblind')
    
set_options()

def read_raster(raster_file):
    """ Read a raster file and return a list of dataframes"""
    ds = rasterio.open(raster_file)
    t = ds.transform
    df_list = []
    # band counts is based 1
    for i in range (1,ds.count+1):
        a = ds.read(i)
        df = pd.DataFrame(a)
        
        # set index and columns to world coordinates
        df.columns = [ (t * (x,0))[0] for x in df.columns]
        df.index = [ ( t * (0,y))[1] for y in df.index]
        
        df_list.append (df)
    ds.close()
    return df_list

def read_raster_binary(source, dtype, width, height, transform):
    ### read raster from binary file , e.g. .bil ###
    raw = np.fromfile(source, dtype=dtype)
    a = raw.reshape(height,width)
    df = pd.DataFrame(a)
    t = affine.Affine(*transform)
    df.columns = [(t*(x,0))[0] for x in df.columns]
    df.index   = [(t*(0,y))[1] for y in df.index]
    return df

def write_raster(df_list, dest_file, color_map=0):
    """ write df raster list to geo tiff together with world file"""
    
    # in case arguent is data frame, not list
    if isinstance(df_list, pd.DataFrame):
        df_list = [df_list]
    
    # create an affine object
    t = calc_affine(df_list[0])

    # get dimension
    bands = len(df_list)
    h, w = df_list[0].shape
    
    # build one 3-dimensional array from the df list
    l = [df.values for df in df_list]
    a = np.array(l)
    dtype = a.dtype

    # write raster, add color_map if defined, TFW = YES: create woldfile!
    with rasterio.open(dest_file, mode='w', driver='GTiff', width=w, height=h, count=bands, dtype=dtype, transform=t, tfw='YES') as dst:
        dst.write(a)
        if color_map:
            dst.write_colormap(1, color_map)

    
def calc_affine(df):
    """generate transorm affine object from raster data frame """

    x0 = df.columns[0]
    y0 = df.index[0]
    dx = df.columns[1] - df.columns[0]
    dy = df.index[1] - df.index[0]
    
    t = affine.Affine(dx, 0, x0 , 0,dy ,y0 ) 
    # x0 + dx because anker point is in the south!
    return t

def vectorize(df):
    """ make shapes from raster, genial! """
    t = calc_affine(df)
    a = df.values
    # zeros an nan are left open space, means mask = True!
    maske = (df != 0).fillna(True)
    gdf = gpd.GeoDataFrame()
    geoms  = []
    value = []
    for s,v in rasterio.features.shapes(a,transform=t,mask=maske.values):
        geoms.append(shape(s))
        value.append(v)
    gdf['geometry'] = geoms
    gdf = gdf.set_geometry('geometry')
    gdf['val']=value
    return gdf

def rasterize(gdf,pixel_size,values=None):
    """ make arry from geo data frame,
    if value = None enumrate shapes starting by 0, empty = -1
    else burn value = (empty_value, burn_value)
    """
    if values==None:
        fill = -1
        geom_value_list = [ (geom,i) for i, geom in enumerate(gdf.geometry)] 
    else:
        fill = values[0]
        geom_value_list = [ (geom,values[1]) for i, geom in enumerate(gdf.geometry)] 
        
        
    x0,y0,x1,y1 = gdf.total_bounds
    
    ulx=pixel_size*int(x0/pixel_size)
    uly=pixel_size*int(y1/pixel_size)+pixel_size
    
    drx=pixel_size*int(x1/pixel_size)+pixel_size
    dry=pixel_size*int(y0/pixel_size)
    
    w,h = int((drx-ulx)/pixel_size), int((uly-dry)/pixel_size)
    
    t = affine.Affine(pixel_size,0,ulx,0,-pixel_size,uly)
    
    result = rasterio.features.rasterize(geom_value_list,out_shape=(h,w),transform=t,fill=fill)
    
    df = pd.DataFrame(result)
    df.columns = [(t*(x,0))[0] for x in df.columns]
    df.index = [(t*(0,y))[1] for y in df.index]
    
    return df
# no unicode characters
def clean(s):
    s=str(s)
    l=[]
    for c in s:
        if ord(c)<256:
            l.append(c)
        else:
            l.append('&#%d;' % ord(c))    
    return ''.join(l)

def refresh_excel(excel_file):
    excel_file=os.path.abspath(excel_file)
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.workbooks.open(excel_file)
    xlapp.Visible = True
    wb.RefreshAll()
    count = wb.Sheets.Count
    for i in range(count):
        ws = wb.Worksheets[i]
        pivotCount = ws.PivotTables().Count
        for j in range(1, pivotCount+1):
            ws.PivotTables(j).PivotCache().Refresh()
    wb.Save()
    xlapp.Quit()


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
                if np.isnan(max_len):
                    max_len=1
                styp='str:%d' % max_len
            props[col]=styp
            
    schema={}
    # set geometry type of the first object for the whole layer
    if len(gdf)>0:
        geo_obj_type=gdf.geometry.iloc[0].geom_type

    else:
        geo_obj_type = 'Point'
        
    schema['geometry']= geo_obj_type
        
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
    with open('mb.mb','w') as fout: 
        fout.write(mb_script)
    os.system('mapbasic.exe -D mb.mb')
    os.system('mapinfow.exe mb.mbx')