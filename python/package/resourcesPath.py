from fbs_runtime.application_context.PySide6 import ApplicationContext

class AppContext:
    _instance = None
    
    @classmethod
    def get(cls):
        if not cls._instance:
            cls._instance = ApplicationContext()
        return cls._instance

