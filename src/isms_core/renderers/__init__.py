from .policy_renderer import PolicyRenderer
from .procedure_renderer import ProcedureRenderer
from .record_renderer import RecordRenderer

def load_renderer(doc_type: str):
    lookup = {
        "Policy": PolicyRenderer,
        "Procedure": ProcedureRenderer,
        "Record": RecordRenderer,
    }
    return lookup.get(doc_type, ProcedureRenderer)
