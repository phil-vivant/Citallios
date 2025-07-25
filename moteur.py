# Standard library imports
import math
import time

# Third party imports
from rich import print
from materia import EC2Concrete, SteelRebar, FibreReinforcedPolymer

from section_flex.geometry.point import Point
from section_flex.geometry.polygon import Polygon
from section_flex.geometry.polygon import Polygon

from section_flex.section.concrete_section import ConcreteSection
from section_flex.section.region import Region
from section_flex.section.rebars import Rebars
from section_flex.section.frp_strips import FRPStrips

# Local applications imports

start_time = time.time()



def unpack_materiaux(materiaux: dict[str: float]):
    fck = materiaux["fck"]
    class_acier = materiaux["class_acier"]
    fyk = materiaux["fyk"]
    Ef = materiaux["Ef"]
    return fck, class_acier, fyk, Ef


def unpack_geometry(geometrie: dict[str: float]):
    h_dalle = geometrie["h_dalle"]
    b_dalle = geometrie["b_dalle"]
    As = geometrie["As"]
    dprim_s = geometrie["dprim_s"]
    ns = geometrie["ns"]
    return h_dalle, b_dalle, As, dprim_s, ns

def unpack_renforts(renforts: dict[str: float]):
    Asr = renforts["Asr"]
    dprim_sr = renforts["dprim_sr"]
    nsr = renforts["nsr"]
    Af = renforts["Af"]
    dprim_f = renforts["dprim_f"]
    nf = renforts["nf"]
    return Asr, dprim_sr, nsr, Af, dprim_f, nf

def unpack_forces(efforts: dict[str: float]):
    m_els_1 = efforts["m_els_1"]
    m_els_2 = efforts["m_els_2"]
    m_elu = efforts["m_elu"]
    m_feu = efforts["m_feu"]
    return m_els_1, m_els_2, m_elu, m_feu


def def_sections(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        comb: str='elu',
):
    # Unpack Data
    fck, class_acier, fyk, Ef = unpack_materiaux(materiaux)
    h_dalle, b_dalle, As, dprim_s, ns = unpack_geometry(geometrie=geometrie)
    Asr, dprim_sr, nsr, Af, dprim_f, nf = unpack_renforts(renforts=renforts)

    # Création des matériaux
    if comb.lower() == 'elu':
        concrete = EC2Concrete(fck=fck, diagram_type="uls_parabola", gamma_c=1.5)
        acier = SteelRebar(ductility_class=class_acier, yield_strength_fyk=fyk, gamma_s=1.15)
        renf_carbone = FibreReinforcedPolymer(modulus_elasticity_ef=Ef)
    if comb.lower() == 'els':
        concrete = EC2Concrete(fck=fck, diagram_type="sls_cracked")
        acier = SteelRebar(ductility_class=class_acier, yield_strength_fyk=fyk, diagram_tpe="sls")
        renf_carbone = FibreReinforcedPolymer(modulus_elasticity_ef=Ef)
    if comb.lower() == 'feu':
        concrete = EC2Concrete(fck=fck, diagram_type="uls_parabola", gamma_c=1)
        acier = SteelRebar(ductility_class=class_acier, yield_strength_fyk=fyk, gamma_s=1)
        renf_carbone = FibreReinforcedPolymer(modulus_elasticity_ef=1e-9)

    # Géométrie Dalle
    pt_00 = Point(-b_dalle / 2, 0.000)
    pt_01 = Point(b_dalle / 2, 0.000)
    pt_02 = Point(b_dalle / 2, h_dalle)
    pt_03 = Point(-b_dalle / 2, h_dalle)

    dalle_poly = Polygon([pt_00, pt_01, pt_02, pt_03])
    dalle_reg = Region(0, [dalle_poly], concrete)

    # Ferraillage de la dalle
    ys = dprim_s
    HA_pts = []
    es = b_dalle / ns
    xs = -es * (ns - 1) / 2
    for i in range(ns):
        HA_pts.append(Point(xs, ys))
        xs += es

    rebar_inf = Rebars(0, As, acier, HA_pts, 21)

    dalle = ConcreteSection(
        regions=[dalle_reg],
        rebars=[rebar_inf],
        frp_strips=[],
        fibre_size_y=b_dalle / 2,
        fibre_size_z=0.005,
    )

    # Acier de renforcement
    rebars=[rebar_inf]
    if Asr * nsr > 0:
        Asr_pts = []
        esr = b_dalle / nsr
        xsr = -esr * (nsr - 1) / 2
        for i in range(nsr):
            Asr_pts.append(Point(xsr, 0))
            xsr += esr
        rebar_renf = Rebars(0, Asr, acier, Asr_pts, 31)
        rebars.append(rebar_renf)

    # FRP de renforcement
    if Af * nf > 0:
        FRP_pts = []
        ef = b_dalle / nf
        xf = -ef * (nf - 1) / 2
        for i in range(nf):
            FRP_pts.append(Point(xf, 0))
            xf += ef
        FRP_inf = FRPStrips(0, Af, renf_carbone, FRP_pts, 1)
    else:
        dummy_frp = FibreReinforcedPolymer(1e-9, 10000, 10000)
        FRP_inf = FRPStrips(0, 0, dummy_frp, [Point(0, 0)], 1)

    dalle_renf = ConcreteSection(
        regions=[dalle_reg],
        rebars=rebars,
        frp_strips=[FRP_inf],
        fibre_size_y=b_dalle / 2,
        fibre_size_z=0.005,
    )

    return dalle, dalle_renf


def print_hypotheses(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
):
    # Unpack Data
    fck, class_acier, fyk, Ef = unpack_materiaux(materiaux)
    h_dalle, b_dalle, As, dprim_s, ns = unpack_geometry(geometrie=geometrie)
    Asr, dprim_sr, nsr, Af, dprim_f, nf = unpack_renforts(renforts=renforts)
    m_els_1, m_els_2, m_elu, m_feu = unpack_forces(efforts=efforts)

    # Présentation Hypothèses
    data = "Rappel des hypothèses :"
    data += f"\n\tSection : \t{round(b_dalle, 2)} m x {round(h_dalle, 2)} m ht"
    data += f"\n\tBéton :\t\tfck = {fck} MPa"
    data += f"\n\tAcier HA :\tfyk = {fyk} MPa"
    data += f"\n\tBarres HA : \t{ns} x As = {round(As * 1e4, 2)} cm²"
    data += f"\n\tRenforts HA : \t{nsr} x Asr = {round(Asr * 1e4, 2)} cm²"
    data += f"\n\tRenforts FRP : \t{nf} x Af = {round(Af * 1e4, 2)} cm²"
    data += f"\n\nRappel des sollicitations :"
    data += f"\n\tM_ELS1 = {round(m_els_1, 2)} kN.m"
    data += f"\n\tM_ELS2 = {round(m_els_2, 2)} kN.m"
    data += f"\n\tM_ELU = {round(m_elu, 2)} kN.m"
    data += f"\n\tM_feu = {round(m_feu, 2)} kN.m"
    print(data)

    return None


def verif_elu(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
):
    section_1, section_2 = def_sections(materiaux, geometrie, renforts, comb='elu')
    ned = 0
    m_els_1, m_els_2, m_elu, m_feu = unpack_forces(efforts=efforts)
    result_elu = "Vérification ELU :"
    result_elu += f"\n\tMoment sollicitant : \tM_ed = {m_elu: .1f} kN.m"
    result_elu += f"\n\tAvant renforcement : \tM_rd1 = {section_1.Mrd_max(ned): .1f} kN.m"
    result_elu += f"\n\tAprès renforcement : \tM_rd2 = {section_2.Mrd_max(ned): .1f} kN.m"
    print(result_elu)

    pod_elu = section_2.from_forces_to_curvature(0, m_elu, 0)
    frp_state_elu = section_2.frp_internal_state(pod_elu)
    sigma_frp_elu = next(iter(frp_state_elu.values()))['stress']
    concrete_state_elu = section_2.concrete_internal_state(pod_elu)
    sigma_conc_elu = 0
    for fibre in concrete_state_elu.values():
        sigma_conc_elu = max(sigma_conc_elu, fibre["stress"])
    rebars_state_elu = section_2.rebars_internal_state(pod_elu)
    sigma_rebar_elu = next(iter(rebars_state_elu.values()))['stress']

    equilibre_elu = "Equilibre de la section renforcée à l'ELU :"
    equilibre_elu += f"\n\tContrainte béton : \tσ_c = {round(sigma_conc_elu, 2)} MPa"
    equilibre_elu += f"\n\tContrainte acier : \tσ_s = {round(sigma_rebar_elu, 1)} MPa"
    equilibre_elu += f"\n\tContrainte carbone : \tσ_f = {round(sigma_frp_elu, 1)} MPa"
    print(equilibre_elu)

    return None


def verif_feu(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
):
    section_1, section_2 = def_sections(materiaux, geometrie, renforts, comb='feu')
    ned = 0
    m_els_1, m_els_2, m_elu, m_feu = unpack_forces(efforts=efforts)
    result_elu = "Vérification au feu :"
    result_elu += f"\n\tMoment sollicitant : \tM_ed = {m_feu: .1f} kN.m"
    result_elu += f"\n\tAvant renforcement : \tM_rd1 = {section_1.Mrd_max(ned): .1f} kN.m"
    result_elu += f"\n\tAprès renforcement : \tM_rd2 = {section_2.Mrd_max(ned): .1f} kN.m"
    print(result_elu)

    pod_feu = section_2.from_forces_to_curvature(0, m_feu, 0)
    frp_state_feu = section_2.frp_internal_state(pod_feu)
    sigma_frp_feu = next(iter(frp_state_feu.values()))['stress']
    concrete_state_feu = section_2.concrete_internal_state(pod_feu)
    sigma_conc_feu = 0
    for fibre in concrete_state_feu.values():
        sigma_conc_feu = max(sigma_conc_feu, fibre["stress"])
    rebars_state_feu = section_2.rebars_internal_state(pod_feu)
    sigma_rebar_feu = next(iter(rebars_state_feu.values()))['stress']

    equilibre_feu = "Equilibre de la section renforcée au feu :"
    equilibre_feu += f"\n\tContrainte béton : \tσ_c = {round(sigma_conc_feu, 2)} MPa"
    equilibre_feu += f"\n\tContrainte acier : \tσ_s = {round(sigma_rebar_feu, 1)} MPa"
    equilibre_feu += f"\n\tContrainte carbone : \tσ_f = {round(sigma_frp_feu, 1)} MPa"
    print(equilibre_feu)

    return None


def verif_els(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
):
    section_1, section_2 = def_sections(materiaux, geometrie, renforts, comb='feu')
    ned = 0
    m_els_1, m_els_2, m_elu, m_feu = unpack_forces(efforts=efforts)

    pod_1 = section_1.from_forces_to_curvature(0, m_els_1, 0)
    pod_2 = section_2.from_forces_to_curvature(0, m_els_2, 0)

    concrete_state_1 = section_1.concrete_internal_state(pod_1)
    sigma_conc_1 = 0
    for fibre in concrete_state_1.values():
        sigma_conc_1 = max(sigma_conc_1, fibre["stress"])
    rebars_state_1 = section_1.rebars_internal_state(pod_1)
    sigma_rebar_1 = next(iter(rebars_state_1.values()))['stress']

    bilan_phase_1 = "Phase 1  - État de contraintes ELS dans la section :"
    bilan_phase_1 += f"\n\tContrainte béton :\t σ_c1 = {round(sigma_conc_1, 2)} MPa"
    bilan_phase_1 += f"\n\tContrainte acier :\t σ_s1 = {round(sigma_rebar_1, 2)} MPa"
    bilan_phase_1 += f"\n\tContrainte carbone :\t σ_f1 = {round(0.00, 2)} MPa"
    print(bilan_phase_1)

    concrete_state_2 = section_2.concrete_internal_state(pod_2)
    sigma_conc_2 = 0
    for fibre in concrete_state_2.values():
        sigma_conc_2 = max(sigma_conc_2, fibre["stress"])
    rebars_state_2 = section_2.rebars_internal_state(pod_2)
    sigma_rebar_2 = next(iter(rebars_state_2.values()))['stress']
    frp_state_2 = section_2.frp_internal_state(pod_2)
    sigma_frp_2 = next(iter(frp_state_2.values()))['stress']

    bilan_phase_2 = "Phase 2  - État de contraintes ELS dans la section :"
    bilan_phase_2 += f"\n\tContrainte béton :\t σ_c2 = {round(sigma_conc_2, 2)} MPa"
    bilan_phase_2 += f"\n\tContrainte acier :\t σ_s2 = {round(sigma_rebar_2, 2)} MPa"
    bilan_phase_2 += f"\n\tContrainte carbone :\t σ_f2 = {round(sigma_frp_2, 2)} MPa"
    print(bilan_phase_2)

    bilan_final = "Bilan  - État de contraintes ELS dans la section :"
    bilan_final += f"\n\tContrainte béton :\t σ_c = {round(sigma_conc_1 + sigma_conc_2, 2)} MPa"
    bilan_final += f"\n\tContrainte acier :\t σ_s = {round(sigma_rebar_1 + sigma_rebar_2, 2)} MPa"
    bilan_final += f"\n\tContrainte carbone :\t σ_f = {round(sigma_frp_2, 2)} MPa"
    print(bilan_final)

    return None

def test():
    diag_no_renf = dalle.build_NM_interaction_diagram(0, 1)
    courbe_add = []
    for ele in diag_no_renf:
        pt = (ele[0], ele[1])
        courbe_add.append(pt)

    # dalle_renf.plot_interaction_diagram_v2(finess=1, add_curves=[courbe_add])
    dalle_renf.plot_geometry_v2()
