class StaticSource:
    GITEE_TOKEN = ""
    GITEE_REPO_OWNER = ""
    GITEE_REPO_NAME = "xionganCleanHubDownload"

    @staticmethod
    def get_current_version() -> str:
        """
        获取当前版本号
        """
        version = "0.0.3"
        return version

    @staticmethod
    def get_software_name() -> str:
        """
        获取软件名称
        """
        name = "雄安清标"
        return name

    @staticmethod
    def get_gitee_token() -> str:
        return StaticSource.GITEE_TOKEN

    @staticmethod
    def get_gitee_repo_owner() -> str:
        return StaticSource.GITEE_REPO_OWNER

    @staticmethod
    def get_gitee_repo_name() -> str:
        return StaticSource.GITEE_REPO_NAME
