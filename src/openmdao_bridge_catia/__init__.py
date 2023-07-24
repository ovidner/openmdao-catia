import enum
import os
from typing import Iterator

import numpy as np
import openmdao.api as om
import scop
from pywintypes import com_error

from .utils import get_catia_session, recast, type_name, update_object


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


# TODO: use a translation regex instead
UNIT_PAIRS = [
    # (CATIA unit, OpenMDAO unit)
    ("m2", "m**2"),
    ("m3", "m**3"),
    ("mm", "mm"),
    ("m", "m"),
    ("N_m2", "N/m**2"),
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


def parameter_type_value_and_unit(param):
    # Dimension params have a lot of subclasses, but they all should have a Unit property
    if hasattr(param, "Unit"):
        type_ = "Dimension"
        unit = param.Unit.Symbol
        value = float(param.Value)
    else:
        unit = None
        type_ = type_name(param)
        if type_ == "RealParam":
            value = float(param.Value)
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


def get_catia_param(root_object, name):
    try:
        return recast(root_object.Parameters.Item(name))
    except com_error:
        raise ValueError(f"Parameter {name} not found in CATIA document.")


def _gen_var_mappings(
    var_dict: dict[str, str | dict | scop.Param], root_object
) -> Iterator[scop.Param]:
    for catia_name, var in var_dict.items():
        if isinstance(var, str):
            given = scop.Param(name=var)
        elif isinstance(var, dict):
            var.pop("val", None)
            given = scop.Param(**var)
        elif isinstance(var, scop.Param):
            given = var
        else:
            raise TypeError(f"Unrecognized variable definition: {var}")

        catia_param = get_catia_param(root_object, catia_name)
        catia_type, catia_val, catia_raw_units = parameter_type_value_and_unit(
            catia_param
        )
        catia_units = units_catia_to_om(catia_raw_units)
        catia_desc = catia_param.Comment

        if given.discrete is None:
            # Provide some reasonable defaults
            discrete = catia_type in ["BoolParam", "IntParam", "StrParam"]
        elif given.discrete is False and catia_type in ["StrParam"]:
            # Stop the user from doing something that will break
            raise ValueError(f"Parameter {catia_name} must be discrete.")
        else:
            # ... and let the user decide otherwise. This implies that
            # also real/dimension values can be used as discrete
            # variables, which might make sense in some cases.
            discrete = given.discrete

        # TODO: notify the user if the units don't match

        meta = given.meta.copy()
        meta["catia-bridge"] = {"name": catia_name}

        yield given.override(
            # FIXME: this should respect the user's default value, but
            # that would need to be unit converted
            default=catia_val,
            # This is the only place where CATIA should get to decide
            units=catia_units or given.units,
            desc=given.desc or catia_desc,
            discrete=discrete,
            meta=meta,
        )


class CatiaComp(om.ExplicitComponent):
    def initialize(self):
        self.options.declare("document", types=(str, os.PathLike))
        self.options.declare("inputs", types=dict)
        self.options.declare("outputs", types=dict)

    def setup(self):
        catia = get_catia_session()
        document = load_document(catia, self.options["document"])
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
            scop.add_input_param(self, input_mapping)

        for output_mapping in self.output_mappings:
            scop.add_output_param(self, output_mapping)

    def compute(self, inputs, outputs, discrete_inputs=None, discrete_outputs=None):
        try:
            for input_mapping in self.input_mappings:
                if input_mapping.discrete:
                    val = discrete_inputs[input_mapping.name]
                else:
                    val = np.ndarray.item(inputs[input_mapping.name])
                input_param = get_catia_param(
                    self.root_object, input_mapping.meta["catia-bridge"]["name"]
                )
                set_parameter_value(
                    input_param,
                    val,
                    input_mapping.units,
                )

            self.root_document.Activate()
            update_object(self.root_object)

            for output_mapping in self.output_mappings:
                output_param = get_catia_param(
                    self.root_object, output_mapping.meta["catia-bridge"]["name"]
                )
                type_, val, catia_unit = parameter_type_value_and_unit(output_param)
                assert catia_unit == units_om_to_catia(output_mapping.units)
                if output_mapping.discrete:
                    discrete_outputs[output_mapping.name] = val
                else:
                    outputs[output_mapping.name] = float(val)
        except com_error as exc:
            raise om.AnalysisError(f"CATIA error: {exc}", exc)
