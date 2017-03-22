import sys,os
import pandas as pd

if 'geopandas' in sys.modules:
    import geopandas as gpd
    
if 'shapely' in sys.modules:
    import shapely

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

WGS="""GEOGCS["unnamed",
    DATUM["WGS_1984",
        SPHEROID["WGS 84",6378137,298.257223563],
        TOWGS84[0,0,0,0,0,0,0]],
    PRIMEM["Greenwich",0],
    UNIT["degree",0.0174532925199433]]
"""

def write_tab(gdf,tab_name,crs_wkt=SWISS):

    def to_multi(row):
        geom=row.geometry
        if geom.type=='Polygon':
            geom=shapely.geometry.MultiPolygon([geom])
        return geom

    
    gdf=gdf.copy()
    
    # make the columns fit for Mapinfo
    new_cols=[]
    for s in gdf.columns:
        s=s.replace(' ','_')
        s=s[0:30]
        new_cols.append(s)
    gdf.columns=new_cols
    
    # make all columns to string, should be improved later
    for a in gdf:
        if a != gdf.geometry.name:
            gdf[a]=gdf[a].astype('str')

    gdf.geometry=gdf.apply(to_multi,axis=1)
    
    # delete files if already there
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

    gdf.to_file(tab_name,driver='MapInfo File',crs_wkt=crs_wkt)    
    
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

def write_mif(df,sMif,x=0,y=0,sCoordSys='swiss'):
    """ Write mif, but only points implemented"""
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

    # wrie mid file
    df.to_csv(sFileTitle+".mid",sep=sSep,header=None,index=None)
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