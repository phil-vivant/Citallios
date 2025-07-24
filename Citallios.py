# Standard library imports

# Third party imports

# Local applications imports
from moteur import calcul_section



materiaux = {
    'fck': 25, 
    'fyk': 500,
    'Ef': 170_000,
}

geom = {
    'h_dalle': 0.25,
    'b_dalle': 2.60,
    'As': 1.13e-4,
    'er_s': 0.04,
    'ns': 4,
    'Af': 0.96e-4,
    'nf': 3,
}

efforts = {
    'm_els_1': 30,
    'm_els_2': 30,
    'm_elu': 80,
}

calcul_section(materiaux=materiaux, geometrie=geom, efforts=efforts)
