# Standard library imports
import math
import time

# Third party imports
import re
from typing import Callable
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
    sigma_fs = materiaux['sigma_fs']
    sigma_fu = materiaux['sigma_fu']
    carbone_feu = materiaux['carbone_feu']
    return fck, class_acier, fyk, Ef, sigma_fs, sigma_fu, carbone_feu


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


def default_id_extractor(k: str) -> int | None:
    m = re.search(r'(\d+)$', k)
    return int(m.group(1)) if m else None


def env_val(val: float | None, default: float=0) -> float:
    if val is None:
        return default
    return float(val)


def envelope(
        d: dict,
        start: int | None = None,
        end: int | None = None,
        field: str = 'stress',
        id_extractor: Callable[[str], int | None] = default_id_extractor
):
    """
    Enveloppe (min/max) du champ `field`.
    - Si `start` et `end` sont None ⇒ toutes les barres/fibres sont prises.
    - Sinon, on filtre par l'ID numérique extrait de la clé via `id_extractor`.
    """
    sub = {}

    for k, v in d.items():
        if field not in v:
            continue

        if start is None and end is None:
            # prendre tout
            sub[k] = v[field]
        else:
            n = id_extractor(k)
            if n is None:
                continue  # clé sans suffixe numérique
            if (start is None or n >= start) and (end is None or n <= end):
                sub[k] = v[field]

    if not sub:
        return {
            'count': 0, 'min': None, 'max': None, 'ids_min': [], 'ids_max': []
        }

    vmin = min(sub.values())
    vmax = max(sub.values())
    return {
        'count': len(sub),
        'min': vmin,
        'max': vmax,
        'ids_min': [k for k, v in sub.items() if v == vmin],
        'ids_max': [k for k, v in sub.items() if v == vmax],
    }



def def_sections(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        comb_type: str='uls',
):
    # Unpack Data
    comb_type = comb_type.lower()
    fck, class_acier, fyk, Ef, sigma_fs, sigma_fu, carbone_feu = unpack_materiaux(materiaux)
    h_dalle, b_dalle, As, dprim_s, ns = unpack_geometry(geometrie=geometrie)
    Asr, dprim_sr, nsr, Af, dprim_f, nf = unpack_renforts(renforts=renforts)

    # Création des matériaux
    if comb_type.lower() == 'uls':
        concrete = EC2Concrete(
            fck=fck, diagram_type="uls_parabola", gamma_c=1.5,
        )
        acier = SteelRebar(
            ductility_class=class_acier, yield_strength_fyk=fyk, gamma_s=1.15,
        )
        renf_carbone = FibreReinforcedPolymer(
            modulus_elasticity_ef=Ef, sigma_f_sls=sigma_fs, sigma_f_uls=sigma_fu
        )
    if comb_type.lower() == 'sls':
        concrete = EC2Concrete(fck=fck, diagram_type="sls_cracked")
        acier = SteelRebar(
            ductility_class=class_acier,
            yield_strength_fyk=fyk,
            diagram_type="sls",
        )
        renf_carbone = FibreReinforcedPolymer(
            modulus_elasticity_ef=Ef, sigma_f_sls=sigma_fs, sigma_f_uls=sigma_fu
        )
    if comb_type.lower() == 'fire':
        concrete = EC2Concrete(fck=fck, diagram_type="uls_parabola", gamma_c=1)
        acier = SteelRebar(
            ductility_class=class_acier, yield_strength_fyk=fyk, gamma_s=1,
        )
        renf_carbone = FibreReinforcedPolymer(modulus_elasticity_ef=Ef*max(carbone_feu, 1e-9))

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

    rebar_inf = Rebars(0, As, acier, HA_pts, 1)

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
        ysr = dprim_sr
        Asr_pts = []
        esr = b_dalle / nsr
        xsr = -esr * (nsr - 1) / 2
        for i in range(nsr):
            Asr_pts.append(Point(xsr, ysr))
            xsr += esr
        rebar_renf = Rebars(0, Asr, acier, Asr_pts, 101)
        rebars.append(rebar_renf)

    # FRP de renforcement
    if Af * nf > 0:
        FRP_pts = []
        yf = dprim_f
        ef = b_dalle / nf
        xf = -ef * (nf - 1) / 2
        for i in range(nf):
            FRP_pts.append(Point(xf, yf))
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
    fck, class_acier, fyk, Ef, sigma_fs, sigma_fu, carbone_feu = unpack_materiaux(materiaux)
    h_dalle, b_dalle, As, dprim_s, ns = unpack_geometry(geometrie=geometrie)
    Asr, dprim_sr, nsr, Af, dprim_f, nf = unpack_renforts(renforts=renforts)
    m_els_1, m_els_2, m_elu, m_feu = unpack_forces(efforts=efforts)

    # Présentation Hypothèses
    data = "Rappel des hypothèses :"
    data += f"\n\tSection :\t{b_dalle:.2f} m x {h_dalle:.2f} m ht"
    data += f"\n\tBéton :\t\tfck = {fck} MPa"
    data += f"\n\tAcier HA :\tfyk = {fyk} MPa\tClasse \"{class_acier}\""
    data += f"\n\tBarres HA : \t{ns} x As = {As * 1e4:.2f} cm²"
    data += f"\td's = {dprim_s:.3f}"
    data += f"\n\tRenforts HA : \t{nsr} x Ar = {Asr * 1e4:.2f} cm²"
    data += f"\td'r = {dprim_sr:.3f}"
    data += f"\n\tRenforts FRP : \t{nf} x Af = {Af * 1e4:.2f} cm²"
    data += f"\td'f = {dprim_f:.3f}"
    data += f"\n\nRappel des sollicitations :"
    data += f"\n\tM_ELS1 = {m_els_1:.1f} kN.m"
    data += f"\n\tM_ELS2 = {m_els_2:.1f} kN.m"
    data += f"\n\tM_ELU  = {m_elu:.1f} kN.m"
    data += f"\n\tM_FEU  = {m_feu:.1f} kN.m"
    print(data)
    return None


def print_results(dic_results: dict[str, float], comb: str='uls'):
    if comb.lower() == 'sls':
        return print_sls_results(sls_results=dic_results)
    else:
        return print_uls_results(uls_results=dic_results, type=comb)


def print_uls_results(uls_results: dict[str, float], type: str="uls") -> str:

    m_ed = uls_results['m_ed']
    m_rd1 = uls_results['m_rd1']
    m_rd2 = uls_results['m_rd2']
    sigma_c = uls_results['sigma_c']
    sigma_s = uls_results['sigma_s']
    sigma_sr = uls_results['sigma_sr']
    sigma_f = uls_results['sigma_f']
    pod_uls = uls_results['pod_uls']

    if type.lower() == 'uls':
        title_1 = f"Vérification ELU :"
        title_2 = f"Équilibre de la section renforcée à l'ELU :"
    elif type.lower() == 'fire':
        title_1 = f"Vérification au feu :"
        title_2 = f"Équilibre de la section renforcée au feu :"

    result_elu = title_1
    result_elu += f"\n\tMoment sollicitant: \tM_Ed  = {m_ed:.1f} kN.m"
    result_elu += f"\n\tAvant renforcement: \tM_Rd1 = {m_rd1:.1f} kN.m"
    result_elu += f"\n\tAprès renforcement: \tM_Rd2 = {m_rd2:.1f} kN.m"

    equilibre_elu = title_2
    equilibre_elu += check_pod_equilibre(pod_uls)
    equilibre_elu += f"\n\tContrainte béton: \tσ_c = {sigma_c:.2f} MPa"
    equilibre_elu += f"\n\tContrainte acier: \tσ_s = {sigma_s:.1f} MPa"
    equilibre_elu += f"\n\tContrainte acier renf: \tσ_r = {sigma_sr:.1f} MPa"
    equilibre_elu += f"\n\tContrainte carbone: \tσ_f = {sigma_f:.1f} MPa"
 
    print(result_elu)
    print(equilibre_elu)

    return None

def print_sls_results(sls_results: dict[str, float]) -> str:

    sigma_c1 = sls_results['sigma_c1']
    sigma_c2 = sls_results['sigma_c2']
    sigma_s1 = sls_results['sigma_s1']
    sigma_s2 = sls_results['sigma_s2']
    sigma_sr2 = sls_results['sigma_sr2']
    sigma_f2 = sls_results['sigma_f2']
    pod_1 = sls_results['pod_1']
    pod_2 = sls_results['pod_2']
    m1 = sls_results['m1']
    m2 = sls_results['m2']

    bilan_phase_1 = "Phase 1  - État de contraintes ELS dans la section :"
    bilan_phase_1 += check_pod_equilibre(pod_1)
    bilan_phase_1 += f"\n\tMoment sollicitant: \tM_1  = {m1:.1f} kN.m"
    bilan_phase_1 += f"\n\tContrainte béton:\tσ_c1 = {sigma_c1:.2f} MPa"
    bilan_phase_1 += f"\n\tContrainte acier:\tσ_s1 = {sigma_s1:.1f} MPa"
    bilan_phase_1 += f"\n\tContrainte acier renf:\tσ_r1 = {0:.1f} MPa"
    bilan_phase_1 += f"\n\tContrainte carbone:\tσ_f1 = {0:.1f} MPa"

    bilan_phase_2 = "Phase 2  - État de contraintes ELS dans la section :"
    bilan_phase_2 += check_pod_equilibre(pod_2)
    bilan_phase_2 += f"\n\tMoment sollicitant: \tM_2  = {m2:.1f} kN.m"
    bilan_phase_2 += f"\n\tContrainte béton:\tσ_c2 = {sigma_c2:.2f} MPa"
    bilan_phase_2 += f"\n\tContrainte acier:\tσ_s2 = {sigma_s2:.1f} MPa"
    bilan_phase_2 += f"\n\tContrainte acier renf:\tσ_r2 = {sigma_sr2:.1f} MPa"
    bilan_phase_2 += f"\n\tContrainte carbone:\tσ_f2 = {sigma_f2:.1f} MPa"

    bilan_final = "Bilan  - État de contraintes ELS dans la section :"
    bilan_final += check_pod_equilibre(pod_1)
    bilan_final += check_pod_equilibre(pod_2)
    bilan_final += f"\n\tMoment ELS: \t\tM_12 = {m1+m2:.1f} kN.m"
    bilan_final += f"\n\tContrainte béton:\tσ_c  = {sigma_c1+sigma_c2:.2f} MPa"
    bilan_final += f"\n\tContrainte acier:\tσ_s  = {sigma_s1+sigma_s2:.1f} MPa"
    bilan_final += f"\n\tContrainte acier renf:\tσ_r  = {sigma_sr2:.1f} MPa"
    bilan_final += f"\n\tContrainte carbone:\tσ_f  = {sigma_f2:.1f} MPa"

    print(bilan_phase_1)
    print(bilan_phase_2)
    print(bilan_final)

    return None


def design_section(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
        comb_type: str = 'uls',
):
    comb_type = comb_type.lower()
    section_1, section_2 = def_sections(
        materiaux, geometrie, renforts, comb_type=comb_type,
    )
    ned = 0
    m_els_1, m_els_2, m_elu, m_fire = unpack_forces(efforts=efforts)

    if comb_type == 'sls':
        pod_1 = section_1.from_forces_to_curvature(ned, m_els_1, 0)
        pod_2 = section_2.from_forces_to_curvature(ned, m_els_2, 0)

        # Phase 1
        concrete_state_1 = section_1.concrete_internal_state(pod_1)
        rebars_state_1 = section_1.rebars_internal_state(pod_1)
        sigma_c1 = envelope(concrete_state_1)['max']
        sigma_s1 = envelope(rebars_state_1, start=1, end=99)['min']


        # Phase 2
        concrete_state_2 = section_2.concrete_internal_state(pod_2)
        rebars_state_2 = section_2.rebars_internal_state(pod_2)
        frp_state_2 = section_2.frp_internal_state(pod_2)
        sigma_c2 = env_val(envelope(concrete_state_2)['max'])
        sigma_s2 = env_val(envelope(rebars_state_2, start=1, end=99)['min'])
        sigma_sr2 = env_val(envelope(rebars_state_2, start=101, end=199)['min'])
        sigma_f2 = env_val(envelope(frp_state_2)['min'])

        acc = {
            'm1': m_els_1,
            'm2': m_els_2,
            'sigma_c1': sigma_c1,
            'sigma_c2': sigma_c2,
            'sigma_s1': sigma_s1,
            'sigma_s2': sigma_s2,
            'sigma_sr2': sigma_sr2,
            'sigma_f2': sigma_f2,
            'pod_1': pod_1,
            'pod_2': pod_2,
        }

        return acc

    if comb_type == 'uls':
        m_ed = m_elu
    elif comb_type == 'fire':
        m_ed = m_fire
    m_rd1 = section_1.Mrd_max(ned)
    m_rd2 = section_2.Mrd_max(ned)

    pod_uls = section_2.from_forces_to_curvature(ned, m_ed, 0)
    concrete_state_uls = section_2.concrete_internal_state(pod_uls)
    rebars_state_uls = section_2.rebars_internal_state(pod_uls)
    frp_state_uls = section_2.frp_internal_state(pod_uls)
    sigma_c = env_val(envelope(concrete_state_uls)['max'])
    sigma_s = env_val(envelope(rebars_state_uls, start=1, end=99)['min'])
    sigma_sr = env_val(envelope(rebars_state_uls, start=101, end=199)['min'])
    sigma_f = env_val(envelope(frp_state_uls)['min'])

    acc = {
        'm_ed': m_ed,
        'm_rd1': m_rd1,
        'm_rd2': m_rd2,
        'sigma_c': sigma_c,
        'sigma_s': sigma_s,
        'sigma_sr': sigma_sr,
        'sigma_f': sigma_f,
        'pod_uls': pod_uls,
    }

    return acc


def verif_els(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
):
    sls_results = design_section(materiaux, geometrie, renforts, efforts, 'sls')
    print_sls_results(sls_results)
    return None


def verif_elu(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
):
    uls_results = design_section(materiaux, geometrie, renforts, efforts, 'uls')
    print_results(uls_results, 'uls')
    return None


def verif_feu(
        materiaux: dict[str: float],
        geometrie: dict[str: float],
        renforts: dict[str: float],
        efforts: dict[str: float],
        
):
    fire_results = design_section(
        materiaux, geometrie, renforts, efforts, 'fire',
    )
    print_results(fire_results, 'fire')
    return None


def check_pod_equilibre(pod) -> str | None:
    equilibre = abs(pod.epsilon_0) + abs(pod.omega_y) + abs(pod.omega_z) != 00
    if not equilibre:
        return f"\n\t!! Warning !! Équilibre non trouvé !"
    return ""
