from openstaad.tools import *
from comtypes import automation
from comtypes import client
from comtypes import CoInitialize

class MaterialExtractor():
    CoInitialize()

    def __init__(self):
        self._staad = client.GetActiveObject("StaadPro.OpenSTAAD")
        self._geometry = self._staad.Geometry
        self._property = self._staad.Property

        # Métodos Geometry
        for m in ["GetMemberCount", "GetBeamList"]:
            self._geometry._FlagAsMethod(m)

        # Métodos Property
        for m in ["GetBeamMaterialName", "GetBeamConstants", "GetBetaAngle"]:
            self._property._FlagAsMethod(m)

    def GetBeamList(self):
        count = self._geometry.GetMemberCount()
        safe_array = make_safe_array_long(count)
        beam_list = make_variant_vt_ref(safe_array, automation.VT_ARRAY | automation.VT_I4)
        self._geometry.GetBeamList(beam_list)
        return beam_list[0]

    def GetBeamMaterialName(self, beam):
        try:
            return self._property.GetBeamMaterialName(beam)
        except Exception:
            return "N/A"

    def GetBeamConstants(self, beam):
        try:
            E = make_variant_vt_ref(make_safe_array_double(1), automation.VT_R8)
            poisson = make_variant_vt_ref(make_safe_array_double(1), automation.VT_R8)
            density = make_variant_vt_ref(make_safe_array_double(1), automation.VT_R8)
            alpha = make_variant_vt_ref(make_safe_array_double(1), automation.VT_R8)
            damp = make_variant_vt_ref(make_safe_array_double(1), automation.VT_R8)

            self._property.GetBeamConstants(beam, E, poisson, density, alpha, damp)

            return {
                "E (kN/mm²)": round(E[0], 4),
                "Poisson": round(poisson[0], 4),
                "Densidad (kg/m³)": round(density[0], 4),
                "Alpha (1/°C)": round(alpha[0], 6),
                "Amortiguamiento": round(damp[0], 4)
            }
        except Exception as e:
            return {"Error": f"No se pudieron obtener propiedades: {e}"}

    def GetBetaAngle(self, beam):
        try:
            return round(self._property.GetBetaAngle(beam), 4)
        except Exception:
            return "N/A"

if __name__ == "__main__":
    extractor = MaterialExtractor()
    beam_list = extractor.GetBeamList()

    print(f"Total de Vigas: {len(beam_list)}\n")

    for beam in beam_list:
        nombre = extractor.GetBeamMaterialName(beam)
        props = extractor.GetBeamConstants(beam)
        beta = extractor.GetBetaAngle(beam)

        print(f"Beam {beam} - Material: {nombre}")
        for key, val in props.items():
            print(f"  {key}: {val}")
        print(f"  Beta Angle: {beta}°")
        print("-" * 40)