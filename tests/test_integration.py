import openmdao.api as om
import pytest
from scop import InnumSpace, IntegerSpace, Param, ParamSet, RealSpace, bool_space

from openmdao_bridge_catia import CatiaComp, utils


@pytest.fixture
def catia_instance():
    print("Starting CATIA...")
    catia = utils.get_catia_session()
    catia.Visible = True
    yield catia
    print("Closing CATIA...")
    catia.Quit()


@pytest.fixture
def part_doc(catia_instance):
    part_doc = utils.recast(catia_instance.Documents.Add("Part"))
    part = part_doc.Part
    params = part.Parameters
    params.CreateReal("Real.1", 0)
    params.CreateInteger("Integer.1", 0)
    params.CreateString("String.1", "0")
    params.CreateBoolean("Boolean.1", False)
    params.CreateDimension("Length.1", "LENGTH", 0)
    params.CreateDimension("Angle.1", "ANGLE", 0)
    params.CreateDimension("Time.1", "TIME", 0)
    params.CreateDimension("Mass.1", "MASS", 0)
    params.CreateDimension("Volume.1", "VOLUME", 0)
    params.CreateDimension("Area.1", "AREA", 0)
    yield part_doc
    part_doc.Close()


@pytest.fixture
def part_path(catia_instance, part_doc, tmpdir):
    part_path = tmpdir / "test_part.CATPart"
    part_doc.SaveAs(str(part_path))
    return part_path


def test_single_eval(catia_instance, part_path):
    prob = om.Problem()
    model = prob.model

    params = ParamSet(
        [
            Param(name="real", space=RealSpace()),
            Param(name="integer", space=IntegerSpace(), discrete=True),
            Param(name="string", space=InnumSpace(), discrete=True),
            Param(name="boolean", space=bool_space(), discrete=True),
            Param(name="length", space=RealSpace(), units="m"),
            Param(name="angle", space=RealSpace(), units="deg"),
            Param(name="time", space=RealSpace(), units="s"),
            Param(name="mass", space=RealSpace(), units="kg"),
            Param(name="volume", space=RealSpace(), units="m**3"),
            Param(name="area", space=RealSpace(), units="m**2"),
        ]
    )

    catia_comp = CatiaComp(
        # instance=catia_instance,
        document=part_path,
        inputs={
            "Real.1": params["real"].override(name="in-real"),
            "Integer.1": params["integer"].override(name="in-integer"),
            "String.1": params["string"].override(name="in-string"),
            "Boolean.1": params["boolean"].override(name="in-boolean"),
            "Length.1": params["length"].override(name="in-length"),
            "Angle.1": params["angle"].override(name="in-angle"),
            "Time.1": params["time"].override(name="in-time"),
            "Mass.1": params["mass"].override(name="in-mass"),
            "Volume.1": params["volume"].override(name="in-volume"),
            "Area.1": params["area"].override(name="in-area"),
        },
        outputs={
            "Real.1": params["real"].override(name="out-real"),
            "Integer.1": params["integer"].override(name="out-integer"),
            "String.1": params["string"].override(name="out-string"),
            "Boolean.1": params["boolean"].override(name="out-boolean"),
            "Length.1": params["length"].override(name="out-length"),
            "Angle.1": params["angle"].override(name="out-angle"),
            "Time.1": params["time"].override(name="out-time"),
            "Mass.1": params["mass"].override(name="out-mass"),
            "Volume.1": params["volume"].override(name="out-volume"),
            "Area.1": params["area"].override(name="out-area"),
        },
    )

    model.add_subsystem("catia_comp", catia_comp)

    try:
        prob.setup()

        prob.set_val("catia_comp.in-real", 1.0)
        prob.set_val("catia_comp.in-integer", 1)
        prob.set_val("catia_comp.in-string", "1")
        prob.set_val("catia_comp.in-boolean", True)
        prob.set_val("catia_comp.in-length", 1.0)
        prob.set_val("catia_comp.in-angle", 1.0)
        prob.set_val("catia_comp.in-time", 1.0)
        prob.set_val("catia_comp.in-mass", 1.0)
        prob.set_val("catia_comp.in-volume", 1.0)
        prob.set_val("catia_comp.in-area", 1.0)

        prob.run_model()

        assert prob.get_val("catia_comp.out-real") == 1.0
        assert prob.get_val("catia_comp.out-integer") == 1
        assert prob.get_val("catia_comp.out-string") == "1"
        assert prob.get_val("catia_comp.out-boolean") == True
        assert prob.get_val("catia_comp.out-length") == 1.0
        assert prob.get_val("catia_comp.out-angle") == 1.0
        assert prob.get_val("catia_comp.out-time") == 1.0
        assert prob.get_val("catia_comp.out-mass") == 1.0
        assert prob.get_val("catia_comp.out-volume") == 1.0
        assert prob.get_val("catia_comp.out-area") == 1.0
    finally:
        prob.cleanup()
