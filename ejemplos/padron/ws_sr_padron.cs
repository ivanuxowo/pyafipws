//<summary>
//Ejemplo de Uso de Interface COM para consultar
//Padron Unico de Contribuyentes AFIP via webservice (servicio web WS-SR-Padron Alcance 4)
//Documentación: http://www.sistemasagiles.com.ar/trac/wiki/PadronContribuyentesAFIP
//2024 (C) Mariano Reingart <reingart@gmail.com>
//Licencia: GPLv3
//</disclaimer>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualBasic;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string wsdl, proxy, cache = "";
            string Path;
            dynamic WSAA = null, Padron = null;
            string tra, cms, ta, certificado, claveprivada;

            Console.WriteLine("DEMO Interfaz PyAfipWs Padron para C#");

            // Crear objeto interface Web Service Autenticación y Autorización
            WSAA = Activator.CreateInstance(Type.GetTypeFromProgID("WSAA"));
            Console.WriteLine(WSAA.Version);

            try
            {
                Console.WriteLine("Generar un Ticket de Requerimiento de Acceso (TRA) para Padron");

                // Especificar la ubicación de los archivos certificado y clave privada
                Path = Environment.CurrentDirectory + "\\";

                certificado = Path + @"..\reingart.crt"; // Certificado de prueba
                claveprivada = Path + @"..\reingart.key"; // Clave privada de prueba

                wsdl = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl";
				ta = WSAA.Autenticar("ws_sr_constancia_inscripcion", certificado, claveprivada, wsdl);

                Console.WriteLine("Ticket de acceso: " + ta);

                // Crear objeto interface Web Service del Padrón
				Padron = Activator.CreateInstance(Type.GetTypeFromProgID("WSSrPadronA5"));
                Console.WriteLine(Padron.Version);
                // Setear token y sign de autorización (pasos previos)
                Padron.Token = WSAA.Token;
                Padron.Sign = WSAA.Sign;

                // CUIT del emisor (debe estar registrado en la AFIP)				
				Padron.Cuit = "20267565393";

                // Conectar al Servicio Web del Padrón
                Padron.Conectar();
                Console.WriteLine("Conectado al Web Service del Padrón");

                // Realizar consultas al padrón
				Console.WriteLine("Ingrese cuit a consultar");
				string cuit = Console.ReadLine();
				object id_persona = (cuit);
                Padron.Consultar(id_persona);
	
                Console.WriteLine("Datos del Padrón para el CUIT: " + id_persona);
				// Imprimir respuesta obtenida
				Console.WriteLine( "Denominacion: " + Padron.denominacion);
				Console.WriteLine( "Tipo: " + Padron.tipo_persona + Padron.tipo_doc + Padron.nro_doc);
				Console.WriteLine("Estado: " + Padron.Estado);
				Console.WriteLine("Direccion: " + Padron.direccion);
				Console.WriteLine("Localidad: " + Padron.localidad);
				Console.WriteLine("Provincia: " + Padron.provincia);
				Console.WriteLine("Codigo Postal: " + Padron.cod_postal);
				
				if (Padron.Excepcion != "" || Padron.Excepcion != null)
				{
					string resultado = $"{Padron.denominacion} {Padron.Estado}\n{Padron.direccion}\n{Padron.localidad}\n{Padron.provincia}\n{Padron.cod_postal}";
					Console.WriteLine(resultado);
					Console.WriteLine(Padron.XmlResponse);
				}
				else
				{
					string rta_afip = $"Error AFIP: {Padron.Excepcion}";
					Console.WriteLine(rta_afip);
					Console.WriteLine(Padron.XmlResponse);
				}
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);

                if (WSAA != null && WSAA.Excepcion != "")
                    Console.WriteLine("Excepción: " + WSAA.Excepcion);

                if (Padron != null && Padron.Traceback != "")
                    Console.WriteLine("Error en Padron: " + Padron.Traceback);
            }

            Console.ReadKey();
        }
    }
}
