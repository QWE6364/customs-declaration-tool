class CustomsToolError(Exception):
    """可直接显示给本地用户的工具异常。"""


class TemplateNotFoundError(CustomsToolError):
    pass


class WorkbookInUseError(CustomsToolError):
    pass


class GenerationError(CustomsToolError):
    pass


class VerificationError(CustomsToolError):
    pass
