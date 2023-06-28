import enum
from pywintypes import com_error
from win32com.client.dynamic import Dispatch as DynamicDispatch
from win32com.client.gencache import EnsureDispatch
from win32com.client import Dispatch, GetObject, GetActiveObject


class SensorType(enum.StrEnum):
    """
    Values to use for the AnalysisLocalSensor.XMLName attribute.

    Source: https://www.eng-tips.com/viewthread.cfm?qid=370727
    """

    # GPS
    DISPLACEMENT_MAGNITUDE = "Sensor_Disp_Iso"
    DISPLACEMENT_VECTOR = "Sensor_Disp"
    RELATIVE_DISPLACEMENT_VECTOR = "Relative_Sensor_Disp"
    ROTATION_VECTOR = "Sensor_Rotation"
    VON_MISES_STRESS = "Sensor_Stress_VonMises"
    ERROR = "Sensor_EstimatedError"

    # EST
    STRESS_TENSOR = "Sensor_Stress_SymTensor"
    PRINCIPAL_SHEARING = "Sensor_Stress_PpalShearing"
    PRINCIPAL_STRESS_TENSOR = "Sensor_Stress_PpalTensor"
    PRINCIPAL_STRAIN_TENSOR = "Sensor_Strain_PpalTensor"
    STRAIN_TENSOR = "Sensor_Strain_SymTensor"
    FORCE = "Sensor_Force"
    MOMENT = "Sensor_Moment"
    ELASTIC_ENERGY = "Sensor_ElasticEnergy"
    CLEARANCE = "Sensor_Display_clearance"
    ACCELERATION_VECTOR = "Sensor_Acceleration"
    RELATIVE_ACCELERATION_VECTOR = "Relative_Sensor_Acceleratio"  # TODO: typo?
    VELOCITY_VECTOR = "Sensor_Velocity"
    RELATIVE_VELOCITY_VECTOR = "Relative_Sensor_Velocity"
    SURFACE_STRESS_TENSOR = "Surface_Sensor_Stress_SymTensor"
    SURFACE_PRINCIPAL_STRESS_TENSOR = "Surface_Sensor_Stress_PpalTensor"


def recast(obj):
    return EnsureDispatch(DynamicDispatch(obj))


def get_catia_session():
    # FIXME: start or get catia
    try:
        return GetObject(Class="CATIA.Application")
    except com_error:
        return Dispatch("CATIA.Application")


def type_name(obj):
    try:
        return getattr(obj, "_oleobj_", obj).GetTypeInfo(0).GetDocumentation(-1)[0]
    except com_error:
        return None


def catia_alive(catia):
    try:
        # Just do something simple that triggers a COM call
        catia.Caption
        return True
    except com_error:
        return False
