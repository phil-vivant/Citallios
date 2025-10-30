# Standard library imports

# Third party imports

# Local applications imports
from moteur import (
    def_sections, print_hypotheses, verif_elu, verif_els, verif_feu,
)


materiaux = {
    'fck': 25,
    'class_acier': "B",
    'fyk': 500,
    'Ef': 170000,
    'sigma_fs': 1300,
    'sigma_fu': 1445,
    'carbone_feu': 1,
}

geometrie = {
    'h_dalle': 0.26,
    'b_dalle': 2.2,
    'As': 5.65e-4,
    'dprim_s': 0.04,
    'ns': 1,
}

renforts = {
    'Asr':0.0e-4,
    'dprim_sr': 0.025,
    'nsr': 10,
    'Af': 0.96e-4,
    'dprim_f': 0.0,
    'nf': 1,
}

efforts = {
    'm_els_1': 38,
    'm_els_2': 0,
    'm_elu': 49.92,
    'm_feu': 38,
}

print_hypotheses(
    materiaux=materiaux, geometrie=geometrie, renforts=renforts, efforts=efforts,
)
verif_els(
    materiaux=materiaux, geometrie=geometrie, renforts=renforts, efforts=efforts,
)
verif_elu(
    materiaux=materiaux, geometrie=geometrie, renforts=renforts, efforts=efforts,
)
verif_feu(
    materiaux=materiaux, geometrie=geometrie, renforts=renforts, efforts=efforts,
)

section_1, section_2 = def_sections(materiaux, geometrie, renforts, 'uls')
section_1.plot_geometry_v2()
section_2.plot_geometry_v2()
section_1.plot_geometry_v2()
section_2.plot_geometry_v2()
section_1_diag = section_1.build_NM_interaction_diagram(theta=0, finess=1)
section_2.plot_interaction_diagram_v2(
    theta=0, finess=1, add_curves=[section_1_diag],
)
