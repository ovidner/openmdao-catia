import pytest
import openmdao.api as om
from openmdao_bridge_catia import utils, CatiaComp
from win32com.client.gencache import EnsureDispatch


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

    catia_comp = CatiaComp(
        # instance=catia_instance,
        file_path=part_path,
        inputs={
            "Real.1": {"name": "in-real", "discrete": False},
            "Integer.1": {"name": "in-integer", "discrete": True},
            "String.1": {"name": "in-string", "discrete": True},
            "Boolean.1": {"name": "in-boolean", "discrete": True},
            "Length.1": {"name": "in-length", "discrete": False},
            "Angle.1": {"name": "in-angle", "discrete": False},
            "Time.1": {"name": "in-time", "discrete": False},
            "Mass.1": {"name": "in-mass", "discrete": False},
            "Volume.1": {"name": "in-volume", "discrete": False},
            "Area.1": {"name": "in-area", "discrete": False},
        },
        outputs={
            "Real.1": {"name": "out-real", "discrete": False},
            "Integer.1": {"name": "out-integer", "discrete": True},
            "String.1": {"name": "out-string", "discrete": True},
            "Boolean.1": {"name": "out-boolean", "discrete": True},
            "Length.1": {"name": "out-length", "discrete": False},
            "Angle.1": {"name": "out-angle", "discrete": False},
            "Time.1": {"name": "out-time", "discrete": False},
            "Mass.1": {"name": "out-mass", "discrete": False},
            "Volume.1": {"name": "out-volume", "discrete": False},
            "Area.1": {"name": "out-area", "discrete": False},
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
