from mailmerge import MailMerge
from datetime import date
template = ("PLANTILLA_CONST_PRUEBA1.docx")
document = MailMerge(template)
print(document.get_merge_fields())
{'numerodeoficio', 'EXPEDIENTE', 'NOMBREalumno', 'MATRICULA', 'CARRERA', 'INGRESOenSEMESTRE', 'DIA', 'MES', 'DIA2', 'MES2', 'AÑO', 'DIA3', 'MES3', 'DIA4', 'MES4', 'DIAS', 'MESexpedicion', 'TITULO', 'NOMBREJefeDepto'}
document . merge ( 
numerodeoficio = '11020',
EXPEDIENTE = 'ETG4-1',
NOMBREalumno = 'RAÚL ESQUIVEL ROSADO',
MATRICULA = 'E17080338',
CARRERA = 'ING. ELECTRÓNICA',
INGRESOenSEMESTRE = '8VO SEMESTRE', 
DIA= '23', 
MES = 'NOVIEMBRE',
DIA2 = '24', 
MES2 ='DICIEMBRE', 
AÑO = '2020',
DIA3 = '12',
MES3 ='ENERO', 
DIA4 = '30',
MES4 = 'OCTUBRE',
DIAS = "20 DIAS",
MESexpedicion = 'OCTUBRE', 
TITULO = 'MSC.',
NOMBREJefeDepto = 'NORA LETICIA CUEVAS CUEVAS')

document.write('CONST-salida.docx')

