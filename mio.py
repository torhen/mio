# v2.0
USE_GEOPANDAS = True

import sys,os,datetime
import pandas as pd
import numpy as np
import subprocess, time, datetime
from PIL import Image
import pathlib
import sqlite3
import json

pd.options.display.max_columns = 200

def now():
    return str(datetime.datetime.now())[0:19]

if USE_GEOPANDAS:
	import geopandas as gpd
	from shapely.geometry import Point, Polygon, MultiPolygon, LineString, MultiLineString, shape
	import rasterio
	import rasterio.features
	import rasterio.mask
	import fiona
	import shapely
	
import matplotlib.pyplot as plt
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor
import win32com.client
import affine

def console(text):
    ts = str(datetime.datetime.now())[0:19]
    b = bytes(f'>> {ts} {text}\n', encoding='latin-1')
    os.write(1, b)

def get_time():
    print(str(datetime.datetime.now())[0:19])

def colset(df, cols_dic):
    """ filter and rename columns in one step"""
    return df[list(cols_dic)].rename(columns=cols_dic)


# cou can get the WKT string unsing ogrinfor file layer
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

WKT_SWISS_95="""\
PROJCS["unnamed",
    GEOGCS["unnamed",
        DATUM["MIF 999,10,674.374,15.156,405.346",
            SPHEROID["Bessel 1841",6377397.155,299.1528128],
            TOWGS84[674.374,15.156,405.346,0,0,0,0]],
        PRIMEM["Greenwich",0],
        UNIT["degree",0.0174532925199433]],
    PROJECTION["Swiss_Oblique_Cylindrical"],
    PARAMETER["latitude_of_center",46.952405555556],
    PARAMETER["central_meridian",7.439583333333],
    PARAMETER["false_easting",2600000],
    PARAMETER["false_northing",1200000],
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

def clean_ascii(df):
    df = df.copy()
    def clean(s):
        s = str(s)
        s = [c for c in s if ord(c) < 256]
        return ''.join(s)

    for col in df.columns:
        if col != 'geometry':
            df[col] = df[col].map(clean)
    return df

def check_path(path):
    path = str(path)
    if os.path.isfile(path):
        title = os.path.basename(path)
        bytes = os.path.getsize(path)
        ctime = datetime.datetime.fromtimestamp(os.path.getctime(path))
        ctime = str(ctime)[0:19]
        size, unit = bytes, 'bytes'
        if bytes > 2**10: size, unit = round(bytes / 2**10, 1), 'KB'
        if bytes > 2**20: size, unit = round(bytes / 2**20, 1), 'MB'
        if bytes > 2**30: size, unit = round(bytes / 2**30, 1), 'GB'
        if bytes > 2**40: size, unit = round(bytes / 2**40, 1), 'TB'
        print(f"ok {title}\t({size} {unit})\tcreated: {ctime}")
        return path
    else:
        abs_path = os.path.abspath(path)
        print(f'MISSING FILE: {abs_path}')
        
def check_folders(*folders):
    ret = True
    for folder in folders:
        folder = str(folder)
        if not os.path.isdir(folder):
            abs_path = os.path.abspath(folder)
            print(f'MISSING FOLDER: {abs_path}')
            ret = False
    return ret

def file_title(path:str):
    return os.path.splitext(os.path.basename(path))[0]

def read_dbf(dbfile):
	"""read dbase file"""
	dbfile = str(dbfile)
	from simpledbf import Dbf5
	dbf = Dbf5(dbfile)
	pl = dbf.to_dataframe()
	pl.columns = [a.split('\x00')[0] for a in pl.columns] # remove strange characters in columns
	return pl

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
	"""to stat notebook from commandline"""
	if len(sys.argv)>1:
		nb = sys.argv[1]
		print('starting',nb)
		run_nb(nb)
	else:
		print('usage: mio.py notebook.jpynb to execute as jupyter notebook')

if __name__ == "__main__":
	main()


def read_raster(raster_file):
	""" Read a raster file and return a list of dataframes"""
	raster_file = str(raster_file) # in case it's a pathlib Path
	ds = rasterio.open(raster_file)
	t = ds.transform
	# changed from t ds.affine
   
	df_list = []
	# band counts is based 1
	for i in range (1,ds.count+1):
		a = ds.read(i)
		df = pd.DataFrame(a)
		
		# set index and columns to world coordinates
		df.columns = [ (t * (x, 0))[0] for x in df.columns]
		df.index =   [ (t * (0 ,y + 1))[1] for y in df.index]
		# y + 1 because of rasterios transformation is for images, not for my dataframe
		# in my datafae I want to keep the lower left corner as index value, not the upper left
		
		df_list.append (df)
	ds.close()
	return df_list

def write_raster(df_list:pd.DataFrame, dest_file, color_map=0):
	dest_file = str(dest_file)
	""" write df raster list to geo tiff together with world file, or Arcview ESRI grid text file
		df  must be 'uint8' to apply color map
		color_map is dictionary like {0:(255,0,0), 1:(0,255,0)}
	"""
	driver_dict = {'.tif': 'GTiff', '.txt': 'AAIGrid', '.asc': 'AAIGrid', '.bil': 'EHdr'}
	driver_string = driver_dict[os.path.splitext(dest_file)[1].lower()]
	
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
	with rasterio.open(dest_file, 
                       mode='w', 
                       driver=driver_string, 
                       width=w, height=h, 
                       count=bands, 
                       dtype=dtype, 
                       transform=t, 
                       tfw='YES',
                       crs =WKT_SWISS
                      ) as dst:
		dst.write(a)
		if color_map:
			dst.write_colormap(1, color_map)

def calc_affine(df):
	"""generate transorm affine object from raster data frame """
	x0 = df.columns[0]
	y0 = df.index[0]
	dx = df.columns[1] - df.columns[0]
	dy = df.index[1] - df.index[0]
	
	t = affine.Affine(dx, 0, x0 , 0, dy ,y0 - dy) 
	# y0 - dy because anker point is in the south!
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

if USE_GEOPANDAS:
	def rasterize(vector_gdf:gpd.GeoDataFrame, raster_df, values_to_burn=128, fill:int=0, all_touched:bool=False):
		""" burn vector features into a raster, input ruster or resolution"""

		# raster_df is integer, create raster with resolution raster_df 
		if isinstance(raster_df, int):
			res = raster_df
			x0, y0, x1, y1 = vector_gdf.geometry.total_bounds
			x0 = int(x0 // res * res - res)
			x1 = int(x1 // res * res + res)
			y0 = int(y0 // res * res - res)
			y1 = int(y1 // res * res + res)
			raster_df = pd.DataFrame(columns=range(x0, x1, res), index=range(y1, y0, -res))

		# no geometry to burn
		if vector_gdf.unary_union.area==0:
			return raster_df
		
		try:
			geom_value_list = zip(vector_gdf.geometry, values_to_burn)
		except:
			geom_value_list = ( (geom, values_to_burn) for geom in vector_gdf.geometry )
		
		t = calc_affine(raster_df)
		
		result = rasterio.features.rasterize(geom_value_list, 
											 out_shape=raster_df.shape, 
											 transform=t, 
											 fill=fill, 
											 all_touched=all_touched)
		
		res_df = pd.DataFrame(result, columns=raster_df.columns, index=raster_df.index)
		return res_df

def refresh_excel(excel_file):
	"""refreshe excel data and pivot tables"""
	excel_file = str(excel_file)
	excel_file = os.path.abspath(excel_file)
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

if USE_GEOPANDAS:
	def write_tab(gdf, tab_name, crs_wkt=WKT_SWISS, index=None):
		"""Write Mapinfo format, all geometry types in one file"""

		gdf.crs = WKT_SWISS
		
		# fiona can't write integer columns, convert columns to float
		for col in gdf.columns:
			stype = str(gdf[col].dtype)
			if stype.startswith('int'):
				gdf[col] = gdf[col].astype(float)
				
		# fiona can't write integer columns, convert index to float
		stype = str(gdf.index.dtype)
		if stype.startswith('int'):
			gdf.index = gdf.index.astype(float)

		gdf.to_file(tab_name,driver='MapInfo File', index=index)    
		return print(len(gdf), 'row(s) written to mapinfo file.')
		
def swiss_wgs(sX,sY):
	"""Aprroximation CH1903 -> WGS84 https://de.wikipedia.org/wiki/Schweizer_Landeskoordinaten"""
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

def wgs_swiss(sLon, sLat):
	"""Aprroximation WGS84 -> CH1903 https://de.wikipedia.org/wiki/Schweizer_Landeskoordinaten"""

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
		
def conv_wgs_swiss (llon, llat, resolution=1):
    lx = []
    ly = []
    for xy in zip(llon, llat):
        x, y = wgs_swiss(*xy)
        x = x // resolution * resolution
        y = y // resolution * resolution
        lx.append(x)
        ly.append(y)
    return lx, ly
    
def run(str_or_list):
	"""Better replacement for os.system()"""
	subprocess.run(str_or_list, check=True, shell=True)	

def run_mb(mb_script:str, mapinfo_path=''):
	"""Run Mapbasic string as mapbasic script
mapinfow.exe and mapbascic : both paths must be set in the PATH env variable!
	"""
	wd = os.getcwd()

	path_mb = os.path.join(wd, 'mb.mb')
	path_mbx = os.path.join(wd, 'mb.mbx')
    
	if os.path.isfile(path_mb): os.remove(path_mb)
	if os.path.isfile(path_mbx): os.remove(path_mbx)

	print(path_mb)
    
	with open(path_mb,'w') as fout: 
		fout.write(mb_script)
    
	# Compile
	subprocess.run(['mapbasic.exe', '-D', path_mb], check=True, shell=True)
	
	# Run
	try:
		subprocess.run([rf'{mapinfo_path}\mapinfow.exe', path_mbx, path_mb], check=True, shell=True)
	except subprocess.CalledProcessError as e:
		print(e)

if USE_GEOPANDAS:
	def disagg(vec:gpd.GeoDataFrame):
		"""Dissagregate collections and multi geomtries"""
		# Split GeometryCollections
		no_coll = []
		for i, row in vec.iterrows():
			geom = row.geometry
			if geom.type == 'GeometryCollection':
				for part in geom:
					row2 = row.copy()
					row2.geometry = part
					no_coll.append(row2)

			else:
					no_coll.append(row)           

		# Split Multi geomries
		res = []
		for row in no_coll:
			geom = row.geometry
			if geom.type.startswith('Multi'):
				for part in geom:
					row2 = row.copy()
					row2.geometry = part
					res.append(row2)
			else:
					res.append(row)

		return gpd.GeoDataFrame(res, crs=vec.crs).reset_index(drop=True)

if USE_GEOPANDAS:
	def write_geojson(vec:gpd.GeoDataFrame, dest):
		"""Write only polygons, including attributes"""
		dest = str(dest)

		# WGS 84
		#vec = vec.to_crs({'init': 'epsg:4326'})

		if os.path.isfile(dest):
			os.remove(dest)
			
		vec.to_file(dest, driver='GeoJSON', encoding='utf-8')


def read_loss(los_file):
    los_file = str(los_file)
    basename = os.path.basename(los_file)
    dbf_path = os.path.dirname(los_file)
    dbf_path = os.path.join(dbf_path, 'pathloss.dbf')
    pl = read_dbf(dbf_path)
    pl = pl.set_index('FILE_NAME')
 
    ser = pl.loc[basename]
    ulx = ser.ULXMAP
    uly = ser.ULYMAP
    nx = ser.NCOLS
    ny = ser.NROWS
    res = ser.RESOLUTION
    #print (basename,ulx,uly,nx,ny,res)
    r = np.fromfile(los_file, dtype='int16')
    r.resize(ny, nx)
    df = pd.DataFrame(r)
    df.columns = np.linspace(ulx, ulx+res*nx-res, nx)
    df.index = np.linspace(uly-res*nx, uly-res, ny)
    df = df.sort_index(ascending=False)
    df = df/16
    return df

def show_perc(i:int, iall:int, istep:int):
    if i % istep == 0:
        print(f'{round(100*i/iall,2)}%', end=' ')
          
def write_sqlite(db_file, df_dict):
    db_conn = sqlite3.connect(db_file)
    for tab_name in df_dict:
        df_dict[tab_name].to_sql(tab_name, db_conn, if_exists="replace", index=None)
    db_conn.close()

def write_json(obj, filename, pretty=True):
    if pretty:
        indent = 4
    else:
        indent = None
    with open(filename, 'w', encoding='utf-8') as fout:
        json.dump(obj, fout,  ensure_ascii=False, indent=indent)
		
def days_per_month(first_day, last_day):
    """return dictionary of month with mumber of days"""
    first_day = pd.to_datetime(first_day)
    last_day = pd.to_datetime(last_day)

    l = pd.date_range(first_day.replace(day=1), last_day.replace(day=1), freq='MS')

    dic = {}
    for e in l:
        dic[str(e)[0:7]] = pd.Period(str(e)).daysinmonth
    
    # patch first and last month
    dic[str(l[0])[0:7]] = dic[str(l[0])[0:7]] - first_day.day + 1
    dic[str(l[-1])[0:7]] = last_day.day
    
    return dic
	
def make_raster_kml(source_file, dest_file):
    cmd = f'gdalwarp -s_srs EPSG:21781 -t_srs EPSG:4326 -overwrite {source_file} {dest_file}'
    os.system(cmd)
            
    ds = rasterio.open(dest_file)
    x0, y0, x1, y1 = ds.bounds
    print(x0, x1, y0, y1)

    s = f"""<kml>
      <Folder>
        <name>Ground Overlays</name>
        <GroundOverlay>
          <name>GSM</name>
          <Icon>
            <href>{dest_file}</href>
          </Icon>
          <LatLonBox>
            <west>{x0}</west>
            <east>{x1}</east>
            <south>{y0}</south>
            <north>{y1}</north>
          </LatLonBox>
        </GroundOverlay>
      </Folder>
    </kml>
    """
    kml_file = pathlib.Path(dest_file).stem + '.kml'
    with open(kml_file, 'w') as fout:
        fout.write(s)

