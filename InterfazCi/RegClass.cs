using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InterfazCi
{
    public class MovPoliza
    {
        public string cuenta;
        public decimal debito;
        public decimal credito;
        public string moneda;
        public string concepto;
        public string referencia;
        public string uuid;
        public string guid;
        public string sn;

    }
    public class Poliza
    {
        public List<MovPoliza> _RegMovtos = new List<MovPoliza>();
        public int Folio;
        public DateTime  FechaAlta;
        public int Ejercicio;
        public int Periodo;
        public int TipoPol;
        public int Clase;
        public int Impresa;
        public string Concepto;
        public decimal Cargos;
        public decimal Abonos;
        public int IdDiario;
        public int SistOrig;
        public int Ajuste;
        public int IdUsuario;
        public int ConFlujo;
        public int ConCuadre;
        public string TimeStamp;
        public string RutaAnexo;
        public string ArchivoAnexo;
        public string Guid;
        public string Referencia;

        public string sMensaje;
    }
    class RegClass
    {
        
    }
}
