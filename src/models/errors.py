class DataNotBesideLabelError(Exception):
    def __init__(self, message='Data is not beside the label! Pass "strict=False" to ignore this'):
        self.message = message
