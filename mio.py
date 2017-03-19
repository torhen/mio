import sys,os
import pandas as pd

if 'geopandas' in sys.modules:
    import geopandas as gpd
    
if 'shaply' in sys.modules:
    from shapely.geometry.multipolygon import MultiPolygon

SWISS="""PROJCS["unnamed",
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

def write_tab(gdf,tab_name):
    def to_multi(row):
        geom=row.geometry
        if geom.type=='Polygon':
            geom=MultiPolygon([geom])
        return geom

    gdf.geometry=gdf.apply(to_multi,axis=1)
    if os.path.isfile(tab_name):
        os.remove(tab_name)
    gdf.to_file(tab_name,driver='MapInfo File',crs_wkt=SWISS)    
    
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