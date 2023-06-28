import dataclasses
import enum
import os
from itertools import chain
from types import SimpleNamespace
from typing import Any, Iterator, Optional
import warnings

import numpy as np
import openmdao.api as om
from pywintypes import com_error
from win32com.client.dynamic import Dispatch as DynamicDispatch
from win32com.client.gencache import EnsureDispatch as Dispatch

from .utils import recast, type_name, get_catia_session

try:
    import scop

    SCOP_AVAILABLE = True
except ImportError:
    scop = SimpleNamespace(Param=SimpleNamespace())
    SCOP_AVAILABLE = False


NOT_SET = object()


@dataclasses.dataclass(frozen=True)
class CatiaVarMapping:
    catia_name: str
    catia_param: Any
    om_name: str
    val: Any = None
    units: str = None
    desc: str = None
    tags: Optional[list] = None
    discrete: bool = False
    shape: tuple = (1,)
    scop_param: scop.Param = None


class RootType(enum.Enum):
    ANALYSIS = enum.auto()
    PART = enum.auto()
    PRODUCT = enum.auto()

    @classmethod
    def from_doc_type_name(cls, type_name):
        if type_name == "AnalysisDocument":
            return cls.ANALYSIS
        elif type_name == "PartDocument":
            return cls.PART
        elif type_name == "ProductDocument":
            return cls.PRODUCT
        else:
            raise ValueError(f"Unrecognized document type: {type_name}")


def coll(collection, *args):
    count = collection.Count
    for i in range(1, count + 1):
        yield collection.Item(i, *args)


INPUT_PARAM_SET_NAME = "OpenMDAO bridge input parameters"
OUTPUT_PARAM_SET_NAME = "OpenMDAO bridge output parameters"

CONTINUOUS_PARAM_TYPES = {"Dimension", "RealParam"}
DISCRETE_PARAM_TYPES = {"BoolParam", "IntParam", "StrParam"}
ALL_PARAM_TYPES = CONTINUOUS_PARAM_TYPES & DISCRETE_PARAM_TYPES

# TODO: use a translation regex instead
UNIT_PAIRS = [
    # (CATIA unit, OpenMDAO unit)
    ("m2", "m**2"),
    ("m3", "m**3"),
    ("mm", "mm"),
    ("m", "m"),
]

CATIA_TO_OM_UNIT_MAP = {catia_unit: om_unit for (catia_unit, om_unit) in UNIT_PAIRS}

OM_TO_CATIA_UNIT_MAP = {om_unit: catia_unit for (catia_unit, om_unit) in UNIT_PAIRS}


def units_catia_to_om(catia_unit):
    return CATIA_TO_OM_UNIT_MAP.get(catia_unit, catia_unit)


def units_om_to_catia(om_unit):
    return OM_TO_CATIA_UNIT_MAP.get(om_unit, om_unit)


def load_document(catia, path, open_=True):
    abs_path = os.path.abspath(path)
    sti_engine = recast(catia.GetItem("CAIEngine"))
    sti_db_item = sti_engine.GetStiDBItemFromCATBSTR(str(abs_path))
    # When the document isn't loaded, it is its own parent in the STI DB. Of course. :|
    if sti_db_item.Parent == sti_db_item:
        old_display_file_alerts = catia.DisplayFileAlerts
        catia.DisplayFileAlerts = False
        try:
            if open_:
                doc = catia.Documents.Open(abs_path)
            else:
                doc = catia.Documents.Read(abs_path)
        finally:
            catia.DisplayFileAlerts = old_display_file_alerts
        return recast(doc)
    else:
        return recast(sti_db_item.GetDocument())


def parameter_unit(param):
    try:
        unit = param.Unit.Symbol
    except (AttributeError, com_error):
        unit = None

    return unit


def parameter_type_value_and_unit(param):
    # Dimension params have a lot of subclasses, but they all should have a Unit property
    if hasattr(param, "Unit"):
        type_ = "Dimension"
        unit = param.Unit.Symbol
        value = float(param.ValueAsString().replace(unit or "", "").strip())
    else:
        unit = None
        type_ = type_name(param)
        if type_ == "RealParam":
            value = float(param.ValueAsString())
        elif type_ in ["BoolParam", "IntParam", "StrParam"]:
            value = param.Value
        else:
            raise ValueError(f"Unrecognized parameter type: {type_}")

    return type_, value, unit


def set_parameter_value(param, val, om_units):
    if hasattr(param, "Unit"):
        type_ = "Dimension"
        catia_units = units_om_to_catia(om_units)
        param.ValuateFromString(f"{val}{catia_units}")
    else:
        type_ = type_name(param)
        if type_ == "RealParam":
            param.ValuateFromString(str(val))
        elif type_ in ["BoolParam", "IntParam", "StrParam"]:
            param.Value = val
        else:
            raise ValueError(f"Unrecognized parameter type: {type_}")


def generate_params(root_object, relatable_parameter_names):
    remaining_params = set(relatable_parameter_names)
    parameters = root_object.Parameters
    for param in coll(parameters):
        if not remaining_params:
            break
        name = parameters.GetNameToUseInRelation(param)
        if name in remaining_params:
            remaining_params.remove(name)
            yield (name, param)
    else:
        if remaining_params:
            raise LookupError(f"The parameters {remaining_params} could not be found.")


def reflect_parameter(source_param, destination_params, destination_param_name):
    source_param_type = type_name(source_param)

    if source_param_type in ["Dimension", "Length", "Angle"]:
        new_param = destination_params.CreateDimension(
            destination_param_name, source_param.Unit.Magnitude, 0
        )
        new_param.ValuateFromString(source_param.ValueAsString())
        return new_param, new_param.Unit.Symbol
    elif source_param_type == "RealParam":
        new_param = destination_params.CreateReal(
            destination_param_name, source_param.Value
        )
        return new_param, None
    else:
        raise TypeError(f"Unknown parameter type {source_param_type}.")


def _gen_var_mappings(var_dict: dict, root_object) -> Iterator[CatiaVarMapping]:
    for catia_name, var in var_dict.items():
        if isinstance(var, str):
            given_om_name = var
            given_discrete = None
            given_val = None
            given_units = None
            given_desc = ""
            given_tags = []
            given_scop_param = None

        elif isinstance(var, dict):
            given_om_name = var["name"]
            given_val = var.get("val", None)
            given_units = var.get("units", None)
            given_discrete = var.get("discrete", None)
            given_desc = var.get("desc", "")
            given_tags = var.get("tags", [])
            given_scop_param = None

        elif SCOP_AVAILABLE and isinstance(var, (scop.Param)):
            given_om_name = var.name
            given_val = var.val
            given_units = var.units
            given_discrete = var.discrete
            given_desc = var.desc
            given_tags = var.tags
            given_scop_param = var

        try:
            catia_param = recast(root_object.Parameters.Item(catia_name))
        except com_error:
            raise ValueError(f"Parameter {catia_name} not found in CATIA document.")
        catia_type, catia_val, catia_raw_units = parameter_type_value_and_unit(
            catia_param
        )
        catia_units = units_catia_to_om(catia_raw_units)
        catia_desc = catia_param.Comment

        if given_discrete is None:
            # Provide some reasonable defaults
            discrete = catia_type in ["BoolParam", "IntParam", "StrParam"]
        elif given_discrete is False and catia_type in ["StrParam"]:
            # Stop the user from doing something that will break
            raise ValueError(f"Parameter {catia_name} must be discrete.")
        else:
            # ... and let the user decide otherwise. This implies that
            # also real/dimension values can be used as discrete
            # variables, which might make sense in some cases.
            discrete = given_discrete

        # TODO: notify the user if the units don't match

        yield CatiaVarMapping(
            catia_name=catia_name,
            catia_param=catia_param,
            om_name=given_om_name,
            val=given_val or catia_val,
            # This is the only place where CATIA gets the upper hand
            units=catia_units or given_units,
            desc=given_desc or catia_desc,
            tags=given_tags,
            discrete=discrete,
            scop_param=given_scop_param,
        )


class CatiaComp(om.ExplicitComponent):
    def initialize(self):
        self.options.declare("file_path", types=os.PathLike)
        self.options.declare("inputs", types=dict)
        self.options.declare("outputs", types=dict)
        self.options.declare("in_parameters", types=dict)
        self.options.declare("out_parameters", types=dict)

    def setup(self):
        catia = get_catia_session()
        document = load_document(catia, self.options["file_path"])
        root_type = RootType.from_doc_type_name(type_name(document))
        if root_type is RootType.ANALYSIS:
            root_object = document.Analysis
        elif root_type is RootType.PART:
            root_object = document.Part
        elif root_type is RootType.PRODUCT:
            root_object = document.Product
        else:
            raise ValueError()

        self.root_document = document
        self.root_object = root_object

        self.input_mappings = list(
            _gen_var_mappings(self.options["inputs"], root_object)
        )
        self.output_mappings = list(
            _gen_var_mappings(self.options["outputs"], root_object)
        )

        for input_mapping in self.input_mappings:
            if input_mapping.discrete:
                self.add_discrete_input(
                    name=input_mapping.om_name,
                    val=input_mapping.val,
                    desc=input_mapping.desc,
                    tags=input_mapping.tags,
                )
            else:
                self.add_input(
                    name=input_mapping.om_name,
                    val=input_mapping.val,
                    units=input_mapping.units,
                    desc=input_mapping.desc,
                    tags=input_mapping.tags,
                )

        for output_mapping in self.output_mappings:
            if output_mapping.discrete:
                self.add_discrete_output(
                    name=output_mapping.om_name,
                    val=output_mapping.val,
                    desc=output_mapping.desc,
                    tags=output_mapping.tags,
                )
            else:
                self.add_output(
                    name=output_mapping.om_name,
                    val=output_mapping.val,
                    units=output_mapping.units,
                    desc=output_mapping.desc,
                    tags=output_mapping.tags,
                )

        # original_params = dict(generate_params(root_object, chain(
        #     self.options["in_parameters"].values(),
        #     self.options["out_parameters"].values(),
        # )))

        # param_sets = root_object.Parameters.RootParameterSet.ParameterSets

        # try:
        #     input_param_set = param_sets.Item(INPUT_PARAM_SET_NAME)
        # except com_error as exc:
        #     if exc.excepinfo[2] == "The method Item failed":
        #         input_param_set = param_sets.CreateSet(INPUT_PARAM_SET_NAME)
        #     else:
        #         raise exc

        # try:
        #     output_param_set = param_sets.Item(OUTPUT_PARAM_SET_NAME)
        # except com_error as exc:
        #     if exc.excepinfo[2] == "The method Item failed":
        #         output_param_set = param_sets.CreateSet(OUTPUT_PARAM_SET_NAME)
        #     else:
        #         raise exc

        # for internal_param_name, catia_param_name in self.options["in_parameters"].items():
        #     original_param = original_params[catia_param_name]
        #     params = input_param_set.AllParameters
        #     rel = original_param.OptionalRelation
        #     if rel:
        #         rel.Parent.Remove(rel.Name)
        #     try:
        #         params.Remove(internal_param_name)
        #     except com_error:
        #         pass
        #     new_param, new_param_unit = reflect_parameter(
        #         original_param,
        #         params,
        #         internal_param_name,
        #     )

        #     root_object.Relations.CreateFormula(f"Input.{internal_param_name}", "", original_param, params.GetNameToUseInRelation(new_param))
        #     self.add_input(internal_param_name, np.nan, units=CATIA_TO_OM_UNIT_MAP[new_param_unit])

        # for internal_param_name, catia_param_name in self.options["out_parameters"].items():
        #     original_param = original_params[catia_param_name]
        #     params = output_param_set.AllParameters
        #     try:
        #         params.Remove(internal_param_name)
        #     except com_error:
        #         pass
        #     new_param, new_param_unit = reflect_parameter(
        #         original_param,
        #         params,
        #         internal_param_name,
        #     )
        #     root_object.Relations.CreateFormula(f"Output.{internal_param_name}", "", new_param, catia_param_name)
        #     self.add_output(internal_param_name, np.nan, units=CATIA_TO_OM_UNIT_MAP[new_param_unit])

        # self.input_params = input_param_set.AllParameters
        # self.output_params = output_param_set.AllParameters

        # self.declare_partials("*", "*", method="fd")

    def compute(self, inputs, outputs, discrete_inputs=None, discrete_outputs=None):
        for input_mapping in self.input_mappings:
            if input_mapping.discrete:
                val = discrete_inputs[input_mapping.om_name]
            else:
                val = np.ndarray.item(inputs[input_mapping.om_name])
            set_parameter_value(input_mapping.catia_param, val, input_mapping.units)

        self.root_document.Activate()
        self.root_object.Update()

        for output_mapping in self.output_mappings:
            type_, val, catia_unit = parameter_type_value_and_unit(
                output_mapping.catia_param
            )
            assert catia_unit == units_om_to_catia(output_mapping.units)
            if output_mapping.discrete:
                discrete_outputs[output_mapping.om_name] = val
            else:
                outputs[output_mapping.om_name] = float(val)

        # for var in self.options["outputs"]:
        #     param = self.root_object.Parameters.Item(var.catia_name)
        #     # We can't really know for sure what unit CATIA will give us, so we'll play it safe.
        #     catia_val, catia_unit = parameter_value_and_unit(param)
        #     val = om.convert_units(
        #         catia_val, CATIA_TO_OM_UNIT_MAP[catia_unit], var.units
        #     )
        #     outputs[var.name] = val
