'''This program connects to database,executes the query,saves the query as kml file and displays it on 
Google Earth'''

# To connect to Postgresql
import psycopg2
import win32com.client
import time

db = psycopg2.connect(dbname="Project",host='localhost',
user='postgres', password='postgres', port=5432)

print 'success'

# Open a cursor to perform database operations
cur=db.cursor()
print 'success'

# Execute a command: this creates a new table
cur.execute("SELECT '<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<kml xmlns=\"http://earth.google.com/kml/2.1\">"
            + "<Document>\n'||textcat_all(replace(replace("
            + "st_askml(geom),'<MultiGeometry>'"
            + ",'<Placemark>'||"
            + "'\n<name>'||'gid '||grid_code||'</name>\n'||"
            + "'<description>'||'gid '||grid_code||'</description>\n'"
            + "||'<MultiGeometry>'),'</MultiGeometry>','</MultiGeometry>"
            + "</Placemark>')||'\n')||'</Document></kml>' FROM LULC2006 where grid_code=4 LIMIT 1")
print 'success'
kmlstring = cur.fetchone()
print 'success'

f = open(r'E:\LULC.kml', 'w')

# print kmlstring[0]
f.write(kmlstring[0])
f.close()

# Close communication with the database
cur.close()
db.close()

# Dispaly Google Earth
ge=win32com.client.Dispatch("GoogleEarth.ApplicationGE")
while not ge.IsInitialized():
    time.sleep(1)
    print "waiting for Google Earth to initialize"
ge.OpenKmlFile(r'E:\LULC.kml',True)
