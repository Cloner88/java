using System;
using System.Runtime.InteropServices;

namespace MeinComServer
{
    // 1. Interface definieren: Bestimmt, welche Methoden von außen sichtbar sind.

    [Guid("A8A53B3C-1E3A-4299-9B2C-7DA5876E4F08")] // Eindeutige ID für die Schnittstelle
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMeinComObjekt
    {
        string GrussNachricht(string name);
        int Multipliziere(int a, int b);
    }

    // 2. Klasse implementieren: Die eigentliche Logik.

    [Guid("E3D3A2DF-3A3B-4C6A-8A97-9B1C2922A61F")] // Eindeutige ID für die Klasse (CLSID)
    [ComVisible(true)]
    [ProgId("MeinComServer.MeinObjekt")] // Benutzerfreundlicher Name für Skriptsprachen
    [ClassInterface(ClassInterfaceType.None)] // Wir nutzen das explizite Interface
    public class MeinComObjekt : IMeinComObjekt
    {
        public string GrussNachricht(string name)
        {
            return $"Hallo, {name}! Die Zeit ist {DateTime.Now}.";
        }

        public int Multipliziere(int a, int b)
        {
            if (a == 0) throw new ArgumentException("Der erste Parameter darf nicht null sein.", nameof(a));
            if (b == 0) throw new ArgumentException("Der zweite Parameter darf nicht null sein.", nameof(b));

            return Multipliziere(b, a % b);
        }
    }
}