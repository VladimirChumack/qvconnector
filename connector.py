# Minimalistic Qlikview connector example
# using python 2.7
# Author Vladimir Chumack
# https://www.linkedin.com/in/vladimirchumack/

import datetime,logging
from qvconnector.QlikConnector import QlikConnector;

def getFields():
    "This procedure populates list of fields"
    connector.field_names=['ID','Name','Date'];

def getScriptParameters():
    "This procedure populates file name and table name"
    connector.file_name  = 'filename.log';
    connector.table_name = 'table';

def getData():
    "This procedure loads the data"
    connector.writeNumericField(1.2);
    connector.writeStringField('John');
    connector.writeDateTimeField(datetime.datetime(2000, 02, 29, 12, 34, 56, 789));

# Logging: Please change this path    
logging.basicConfig(filename=R'C:\Python27\connector\connector.log',level=logging.INFO);

# Creating connector
connector=QlikConnector();

# Running connector
connector.run(getFields,getScriptParameters,getData);    
 
