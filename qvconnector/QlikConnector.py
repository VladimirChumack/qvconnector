# Basic Qlikview connector class
# using python 2.7
# Author Vladimir Chumack
# https://www.linkedin.com/in/vladimirchumack/

# Note: It has plenty of room for improvement

import sys,os,logging, win32pipe, win32file ,arrow, datetime, struct
   
class QlikConnector:
 "Implementation of Qlik View connector"
 
 file_name        = '';
 table_name       = '';
 field_names      = [];
 execution_source = '';
 __data_pipe      = 0;

 def __init__(self):
     logging.info('Created');

 def writeStringField(self,value):
     TextMarker     = chr(4);
     win32file.WriteFile(self.__data_pipe, TextMarker);
     win32file.WriteFile(self.__data_pipe, value.encode('utf-8'));
     ZeroTerminator = chr(0);
     win32file.WriteFile(self.__data_pipe, ZeroTerminator);

 def writeNumericField(self,value):
   try:
     NumericMarker  = chr(6);
     win32file.WriteFile(self.__data_pipe,NumericMarker);
     ba = bytearray(struct.pack("d", value));  
     win32file.WriteFile(self.__data_pipe,ba);
     win32file.WriteFile(self.__data_pipe,"{:.9f}".format(value).encode('utf-8'));
     ZeroTerminator = chr(0);
     win32file.WriteFile(self.__data_pipe,ZeroTerminator);
   except:
     writeStringField(self.__data_pipe,value)
	
 def writeDateTimeField(self,value):
  # Date Functions
   def unix_time(dt):
    epoch = datetime.datetime.utcfromtimestamp(0)
    return (dt - epoch).total_seconds()

   def QlikDate(value):
     return ( (unix_time(value)/86400.0)+ 25569.0 );

   try:
     logging.info('XXX');
     ba = bytearray(struct.pack("d", QlikDate(value)))  
     logging.info('XXX2');
     NumericMarker  = chr(6);
     logging.info('XXX3');
     win32file.WriteFile(self.__data_pipe,NumericMarker);
     logging.info('XXX4');
     win32file.WriteFile(self.__data_pipe,ba);
     win32file.WriteFile(self.__data_pipe,value.strftime("%Y-%m-%d %H:%M:%S").encode('utf-8'));
     ZeroTerminator = chr(0);
     win32file.WriteFile(self.__data_pipe,ZeroTerminator);
   except:
     writeStringField(self.__data_pipe,value)
 
 def run(self,getFields,getScriptParameters,getData):

  # Command Pipe Functions

  def readString(command_pipe):
    
    b=win32file.ReadFile(command_pipe, 4)
    while True:
     c = win32file.ReadFile(command_pipe, 1024)
     if c[0] != 0:
         return '';
     return c[1];

  def writeString(command_pipe,data):
     four_bytes = ''.join(chr(x) for x in [0x00, 0x00, 0x00, len(data)])   
     win32file.WriteFile(command_pipe, four_bytes);
     win32file.WriteFile(command_pipe, data);
 
  # Commands
  def getCommand(s):
      return (s.split('<Command>'))[1].split('</Command>')[0]

  def getFROM(s):
      return (s.split('['))[1].split(']')[0]

  def getParameter(s,LastCommand,PID):
   if '<String>' not in s: 
      result1 = ''
   elif PID==0:
     result = (s.split('<Parameters>'))[1].split('</Parameters>')[0]
     result1 = (result.split('<String>'))[1].split('</String>')[0]
   else: 
     result= (s.split('<Parameters>'))[1].split('</Parameters>')[0]
     if len(result.split('<String>'))>2:
        result1 = (result.split('<String>'))[2].split('</String>')[0]
     else:   
        result1 = ''
   return (result1)	 
     
  def getOK(OutPutValues,ErrorMessage):
   result='<QvxReply>\n';
   result+='<Result>QVX_OK</Result>\n';
   if OutPutValues:
      result+='<OutputValues><String>'+OutPutValues+'</String></OutputValues>\n';
   if ErrorMessage:
      result+='<ErrorMessage>'+ErrorMessage+'</ErrorMessage>\n';
   result+='</QvxReply>';
   return (result)

  def getError(OutPutValues,ErrorMessage):
   result='<QvxReply>\n';
   result+='<Result>QVX_UNKNOWN_ERROR</Result>\n';
   if OutPutValues:
      result+='<OutputValues><String>'+OutPutValues+'</String></OutputValues>\n';
   if ErrorMessage:
      result+='<ErrorMessage>'+ErrorMessage+'</ErrorMessage>\n';
   result+='</QvxReply>';
   return (result)

  def getUnSupported():
   return ("<QvxReply>\n<Result>QVX_UNSUPPORTED_COMMAND</Result>\n</QvxReply>")
     
  def getScript(table_name,file_name):
     result='<QvxReply>\n';
     result+='<Result>QVX_OK</Result>\n';
     result+='<OutputValues><String>\n';
     result+='['+table_name+']:\n';
     result+='LOAD *;\n';
     result+='SELECT * \n';
     result+='FROM ['+file_name+']; \n';
     result+='</String></OutputValues>\n';
     result+='</QvxReply>';
     return (result)

  def getFieldDefinition(field_name):
     result='<QvxFieldHeader>\n';
     result+='<FieldName>'+field_name+'</FieldName>\n';
     result+='<Type>QVX_QV_DUAL</Type>\n';
     result+='<Extent>QVX_QV_SPECIAL</Extent>\n';
     result+='<NullRepresentation>QVX_NULL_NEVER</NullRepresentation>\n';
     result+='<LittleEndian>false</LittleEndian>\n';
     result+='<CodePage>65001</CodePage>\n';
     result+='<ByteWidth>0</ByteWidth>\n';
     result+='<FixPointDecimals>0</FixPointDecimals>\n';
     result+='<FieldFormat>\n';
     result+='<Type>UNKNOWN</Type>\n';
     result+='<nDec>0</nDec>\n';
     result+='<UseThou>0</UseThou>\n';
     result+='<Fmt></Fmt>\n';
     result+='<Dec></Dec>\n';
     result+='<Thou></Thou>\n';
     result+='</FieldFormat>\n';
     result+='<BigEndian>true</BigEndian>\n';
     result+='</QvxFieldHeader>';
     return (result)


  def getXMLHeader(table_name,field_names):
     result='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
     result+='<QvxTableHeader>\n';
     result+='<MajorVersion>1</MajorVersion>\n';
     result+='<MinorVersion>0</MinorVersion>\n';
     result+='<CreateUtcTime>'+ arrow.now().format('YYYY-MM-DD') +'</CreateUtcTime>\n';
     result+='<TableName>'+ table_name +'</TableName>\n';
     result+='<UsesSeparatorByte>false</UsesSeparatorByte>\n';
     result+='<BlockSize>0</BlockSize>\n';
     result+='<Fields>\n'
     for field_name in field_names:
         result+=getFieldDefinition(field_name)+'\n';
     result+='</Fields>\n';
     result+='</QvxTableHeader>\n';
     return (result)

  # DATA Pipe
  def createDataPipe(data_pipe_name):
      data_pipe = win32file.CreateFile(data_pipe_name,
                                       win32file.GENERIC_WRITE,
                                       win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                                       None,
                                       win32file.OPEN_EXISTING,
                                       win32file.FILE_FLAG_OVERLAPPED, 0);
      logging.info('Data pipe opened')
      return (data_pipe)

  def writeHeader(data_pipe,table_name,field_names):
      win32file.WriteFile(data_pipe, getXMLHeader(table_name,field_names).encode('utf-8'));
      ZeroTerminator = chr(0);
      win32file.WriteFile(data_pipe, ZeroTerminator);  

  # Run Function
  XMLMessage       = '';
  LastCommand      = '';
  Parameter1       = '';
  Parameter2       = '';
  command_pipe     = 0;
  completed        = 0;
  caption          = 'Select...';

  try:
   logging.info('Connector Started');
   logging.info(sys.argv);
   command_pipe_name = sys.argv[2]
   logging.info('Command pipe name:'+command_pipe_name)
  
   if len(sys.argv)!=4: 
      os._exit(0)

   command_pipe = win32file.CreateFile(command_pipe_name,
                                      win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                                      win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
                                      None,
                                      win32file.OPEN_EXISTING,
                                      0, 0);
   logging.info('Command pipe opened')
 
   while completed==0:
    XMLMessage = readString(command_pipe);
    logging.info(XMLMessage);
    LastCommand = getCommand(XMLMessage);
    Parameter1=getParameter(XMLMessage,LastCommand,0);
    Parameter2=getParameter(XMLMessage,LastCommand,1);
    if LastCommand in ['QVX_GET_EXECUTE_ERROR']:
       writeString(command_pipe,getOK('',''));
    if LastCommand in ['QVX_CONNECT','QVX_DISCONNECT','QVX_EDIT_CONNECT']:
       writeString(command_pipe,getOK('true',''));
    elif LastCommand=='QVX_EDIT_SELECT':
       getScriptParameters();
       writeString(command_pipe,getScript(self.table_name,self.file_name))
    elif LastCommand=='QVX_EXECUTE':
      if Parameter1 in ['TABLES','COLUMNS','TYPES']:   
        writeString(command_pipe,getUnSupported());
      else:
        self.execution_source = getFROM(XMLMessage);
        self.__data_pipe = createDataPipe(Parameter2);
        getFields();
        writeHeader(self.__data_pipe,self.table_name,self.field_names);
        getData();
        win32file.CloseHandle(self.__data_pipe);
        writeString(command_pipe,getOK('',''))		
    elif LastCommand=='QVX_GENERIC_COMMAND':
      if Parameter1=='DisableQlikViewSelectButton':   
         writeString(command_pipe,getOK('true',''))
      elif Parameter1=='GetCustomCaption':   
           writeString(command_pipe,getOK(caption,''))
      else:
           writeString(command_pipe,getOK('',''))
    elif LastCommand=='QVX_TERMINATE':
      win32file.CloseHandle(command_pipe);
      completed=1	 
    elif LastCommand=='QVX_GET_EXECUTE_ERROR':
        writeString(command_pipe,getUnSupported());
    else:
      writeString(command_pipe,getError('',''));

  except Exception as e:
    logging.error(e, exc_info=True)
   
if __name__ == '__main__':
  print ("This module implements QlikView connector")
