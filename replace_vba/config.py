class Config:
    def index(self):
        self.target_file_path = None

    def set_target_path(self, path):
        self.target_file_path = path

    def get_target_path(self):
        return self.target_file_path

config = Config()

