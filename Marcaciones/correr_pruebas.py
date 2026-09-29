from __future__ import annotations

import os
import sys
import unittest

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
sys.path.insert(0, os.path.join(BASE, "pruebas"))


def main() -> int:
    print("=" * 74)
    print("PRUEBAS UNITARIAS")
    print("=" * 74)
    cargador = unittest.TestLoader()
    carpeta_pruebas = os.path.join(BASE, "pruebas")
    suite = cargador.discover(carpeta_pruebas, pattern="test_*.py", top_level_dir=carpeta_pruebas)
    resultado = unittest.TextTestRunner(verbosity=2).run(suite)
    exito_unit = resultado.wasSuccessful()

    if "--solo-unit" in sys.argv:
        return 0 if exito_unit else 1

    print("\n" + "=" * 74)
    import estres
    exito_estres = estres.main()
    return 0 if (exito_unit and exito_estres) else 1


if __name__ == "__main__":
    sys.exit(main())
