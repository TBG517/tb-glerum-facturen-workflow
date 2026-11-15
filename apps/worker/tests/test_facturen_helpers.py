"""Lokale test voor de pure hulpfuncties van de facturenworkflow.

De Mistral Workflows SDK is niet nodig: mistralai.workflows wordt gestubd
voordat de module geimporteerd wordt. Draaien met:

    python3 test_facturen_helpers.py
"""

import sys
import types


def _stub_sdk() -> None:
    if "mistralai.workflows" in sys.modules:
        return
    mistralai = types.ModuleType("mistralai")
    workflows = types.ModuleType("mistralai.workflows")
    plugins = types.ModuleType("mistralai.workflows.plugins")
    mistralai_plugins = types.ModuleType("mistralai.workflows.plugins.mistralai")

    class _WorkflowError(Exception):
        def __init__(self, message: str, non_retryable: bool = False):
            super().__init__(message)
            self.non_retryable = non_retryable

    workflows.WorkflowError = _WorkflowError
    workflows.activity = lambda **kwargs: (lambda fn: fn)
    workflows.workflow = types.SimpleNamespace(
        define=lambda **kwargs: (lambda cls: cls),
        entrypoint=lambda fn: fn,
    )

    mistralai_plugins.ChatCompletionRequest = object
    mistralai_plugins.DocumentURLChunk = object
    mistralai_plugins.OCRRequest = object
    mistralai_plugins.UserMessage = object
    mistralai_plugins.chat_parse_to_model = None
    mistralai_plugins.get_mistral_client = lambda: None
    mistralai_plugins.mistralai_ocr = None

    mistralai.workflows = workflows
    mistralai.workflows.plugins = plugins
    plugins.mistralai = mistralai_plugins
    sys.modules["mistralai"] = mistralai
    sys.modules["mistralai.workflows"] = workflows
    sys.modules["mistralai.workflows.plugins"] = plugins
    sys.modules["mistralai.workflows.plugins.mistralai"] = mistralai_plugins

    pydantic_stub = types.ModuleType("pydantic")


    class _Base:
        def __init__(self, **data):
            for k, v in data.items():
                setattr(self, k, v)


    pydantic_stub.BaseModel = _Base
    sys.modules.setdefault("pydantic", pydantic_stub)


def main() -> None:
    _stub_sdk()
    sys.path.insert(0, "src/worker/workflows")
    import facturen  # noqa: E402

    # _fmt_date
    assert facturen._fmt_date("2026-10-06T00:00:00Z", "") == "06-10-2026", "ISO-datum"
    assert facturen._fmt_date("06-10-2026", "") == "06-10-2026", "NL-datum"
    assert facturen._fmt_date("", "2026-10-09T18:32:58Z") == "09-10-2026", "fallback"
    assert facturen._fmt_date("geen datum", "ook niet") == "", "geen match"

    # _filename: JJMMDD Naam fact. nummer.ext
    row = facturen.InvoiceRow(
        datum="06-10-2026",
        naam="MoveMove",
        categorie="Auto",
        factuur_nummer="12773558",
    )
    assert facturen._filename(row, ".pdf") == "261006 MoveMove fact. 12773558.pdf"
    row2 = facturen.InvoiceRow(datum="09-10-2026", naam='Test "kwoting"', categorie="Overig", factuur_nummer="")
    name2 = facturen._filename(row2, ".pdf")
    assert name2 == '261009 Test kwoting fact. onbekend.pdf', name2

    # excel-rijindeling: 8 kolommen in de vaste volgorde
    excel = row.excel_row()
    assert excel == ["06-10-2026", "MoveMove", "Auto", "12773558", "", "", "", ""]
    assert len(excel) == 8

    print("ALLE TESTS GESLAAGD")


if __name__ == "__main__":
    main()
