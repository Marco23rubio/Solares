import math;
import pandas as pd;

def calcular_hi(n_dias,inclinacion,irradiacion_diaria_horizontal):
    valor360=360;
    dias_anual=365;
    latitud=24.844195;
    Gcs=1361;

    S=round(math.sin(math.radians(valor360*(284+n_dias)/dias_anual))*23.45,4);
    Ws=round(math.degrees(math.acos(-math.tan(latitud * math.pi/180) * math.tan(S * math.pi/180))),4);
    HohJ_m=round(((24*3600)/math.pi)*Gcs*(1+0.033*math.cos(math.radians((valor360*n_dias)/dias_anual)))*(math.cos(math.radians(latitud))*math.cos(math.radians(S))*math.sin(math.radians(Ws))+((math.pi/180)*Ws)*math.sin(math.radians(latitud))*math.sin(math.radians(S))),4);
    Hoh=round(HohJ_m/1e6,4);
    Kt=round(irradiacion_diaria_horizontal/Hoh,4);
    Fd=round(1+(0.2832*Kt)-(2.5557*Kt**2)+(0.8448*Kt**3),4);
    Hdh=round(Fd*irradiacion_diaria_horizontal,4);
    Hbh=round(irradiacion_diaria_horizontal-Hdh,4);
    Ot=round(inclinacion-abs(latitud),4);
    Rb=round((((math.pi*Ws)/180)*(math.sin(math.radians(S))*math.sin(math.radians(Ot)))+(math.cos(math.radians(S))*math.cos(math.radians(Ot))*math.sin(math.radians(Ws))))/(((math.pi*Ws)/180)*(math.sin(math.radians(S))*math.sin(math.radians(latitud)))+(math.cos(math.radians(S))*math.cos(math.radians(latitud))*math.sin(math.radians(Ws)))),4);
    Hbi=round(Rb*Hbh,4);
    Hdic=round(Hdh*(1+math.cos(math.radians(inclinacion)))/2,4);
    Hdir=round((irradiacion_diaria_horizontal*0.6*(1-math.cos(math.radians(35))))/2,4);
    Hi=round(Hbi+Hdic+Hdir,4);
    

    return Hi;

def calculo_dias():
    try:
        excel_data_df=pd.read_excel('solares-back\Solar.xlsm',sheet_name='irradicion');
        data = {'n_dias': [], 'inclinacion': [], 'irradiacion': [], 'Hi': []};
        x=0;
        for n_dias in range(1,366):
            irradiacion=excel_data_df['irradiacion'].tolist()[x];
            irradiacion=round(float(irradiacion),4);
            x +=1;
            for inclinacion in range(0, 91, 10):
                Hi = calcular_hi(n_dias, inclinacion,irradiacion);
                data['n_dias'].append(n_dias);
                data['inclinacion'].append(inclinacion);
                data['irradiacion'].append(irradiacion);
                data['Hi'].append(Hi);
        df = pd.DataFrame(data);
        df.to_excel('resultados.xlsx',index=False);
        print('El programa se ejecutó de manera exitosa,revisa tu archivo "resultados.xlsx" para conocer tus irradiaciones');
    except PermissionError as error:
        print(error);
        print('UPS!, parece que hubo un error, intenta cerrar tu excel "resultados.xlsx" y ejecuta de nuevo el programa :D');
    except IndexError as error:
        print(error);
        print("UPS!, parece que tu rango en python recorre por más dias que tienes en Excel, entonces agrega más datos a Excel o reduce tu rango de n_dias en Python :3");
    
# sys.stdout = open("Irradiaciones.txt","w");

if __name__ == '__main__':
    calculo_dias();

    

