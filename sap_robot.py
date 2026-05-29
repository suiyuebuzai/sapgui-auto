import win32com.client
import time
import yaml
from pathlib import Path
from typing import List, Dict, Any, Optional


class SAPRobot:
    """SAP GUI Scripting API 封装，提供原子操作方法。"""

    def __init__(self, system_name: str = None, timeout: int = 10):
        self.timeout = timeout
        self.session = None
        self._connect(system_name)

    def _connect(self, system_name: str = None) -> bool:
        """
        连接 SAP GUI。优先复用已有会话，没有则自动登录。
        - 已有会话 → 直接复用
        - 没有会话 → 读取 sap_connections.yaml → 自动登录
        """
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine

            # 已有连接，直接复用
            if application.Children.Count > 0:
                connection = application.Children(0)
                if connection.Children.Count > 0:
                    self.session = connection.Children(0)
                    return True

            # 没有会话，走自动登录
            if system_name is None:
                system_name = self._get_default_system()
            return self._login(system_name)
        except Exception as e:
            print(f"连接 SAP GUI 失败：{e}")
            print("请确保 SAP GUI 已打开，且已启用 Scripting")
            return False

    def _get_default_system(self) -> str:
        """从 sap_connections.yaml 读取默认系统名。"""
        config_path = Path("sap_connections.yaml")
        if not config_path.exists():
            raise FileNotFoundError(
                "sap_connections.yaml 不存在，请创建凭据配置文件后重试"
            )
        config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        return config.get("default", "")

    def _login(self, system_name: str) -> bool:
        """读取凭据并登录指定 SAP 系统。"""
        config_path = Path("sap_connections.yaml")
        if not config_path.exists():
            print("sap_connections.yaml 不存在，无法自动登录")
            return False
        config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        conn = next(
            (c for c in config.get("connections", []) if c["name"] == system_name),
            None,
        )
        if not conn:
            names = [c["name"] for c in config.get("connections", [])]
            print(f"系统 '{system_name}' 未在配置中找到。可用系统：{names}")
            return False

        # 执行登录
        if not self.open_connection(system_name):
            return False
        self.set_text("wnd[0]/usr/txtRSYST-MANDT", conn.get("client", ""))
        self.set_text("wnd[0]/usr/txtRSYST-BNAME", conn.get("user", ""))
        self.set_text("wnd[0]/usr/pwdRSYST-BCODE", conn.get("password", ""))
        if conn.get("language"):
            self.set_text("wnd[0]/usr/txtRSYST-LANGU", conn["language"])
        self.press_key("ENTER")
        time.sleep(2)

        # 检查登录结果
        status = self.get_status()
        if status and ("错误" in status or "Error" in status or "invalid" in status.lower()):
            print(f"登录失败：{status}")
            return False
        return True

    def open_connection(self, system_name: str) -> bool:
        """通过 SAP Logon Pad 打开系统连接。"""
        try:
            application = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            connection = application.OpenConnection(system_name, True)
            self.session = connection.Children(0)
            time.sleep(2)
            return True
        except Exception as e:
            print(f"打开连接 '{system_name}' 失败：{e}")
            return False

    def open_tcode(self, tcode: str) -> bool:
        """打开事务码。"""
        try:
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.Text = tcode
            self.session.findById("wnd[0]").SendVKey(0)
            time.sleep(1)
            return True
        except Exception as e:
            print(f"打开事务码 {tcode} 失败：{e}")
            return False

    def set_text(self, field_id: str, value: str) -> bool:
        """填写文本字段。"""
        try:
            field = self.session.findById(field_id)
            field.Text = str(value)
            return True
        except Exception as e:
            print(f"填写字段 {field_id} 失败：{e}")
            return False

    def press_key(self, key: str) -> bool:
        """按功能键。key 取值：ENTER / F3 / F8 / F12 / CTRL+S"""
        key_map = {"ENTER": 0, "F3": 3, "F8": 8, "F12": 12, "CTRL+S": 26}
        try:
            vkey = key_map.get(key.upper(), 0)
            self.session.findById("wnd[0]").SendVKey(vkey)
            time.sleep(0.5)
            return True
        except Exception as e:
            print(f"按键 {key} 失败：{e}")
            return False

    def click_button(self, button_id: str) -> bool:
        """点击按钮。"""
        try:
            self.session.findById(button_id).Press()
            return True
        except Exception as e:
            print(f"点击按钮 {button_id} 失败：{e}")
            return False

    def get_table_data(self, table_id: str, max_rows: int = 100) -> List[List[str]]:
        """
        读取 GridView 表格数据，返回行列二维数组。
        列名通过 grid.ColumnOrder 获取（SAP GridView.GetCellValue 要求列名字符串，非整数索引）。
        """
        try:
            grid = self.session.findById(table_id)
            rows = min(grid.RowCount, max_rows)
            col_names = list(grid.ColumnOrder)
            result = []
            for row in range(rows):
                row_data = []
                for col_name in col_names:
                    try:
                        row_data.append(grid.GetCellValue(row, col_name))
                    except Exception:
                        row_data.append("")
                result.append(row_data)
            return result
        except Exception as e:
            print(f"读取表格 {table_id} 失败：{e}")
            return []

    def get_status(self) -> str:
        """读取状态栏消息。"""
        try:
            return self.session.findById("wnd[0]/sbar").Text
        except Exception:
            return ""

    def wait_for_screen(self, screen_id: str = "wnd[0]", timeout: Optional[int] = None) -> bool:
        """等待屏幕加载完成。"""
        deadline = time.time() + (timeout or self.timeout)
        while time.time() < deadline:
            try:
                self.session.findById(screen_id)
                return True
            except Exception:
                time.sleep(0.5)
        return False

    def execute_sequence(
        self,
        steps: List[Dict[str, Any]],
        arguments: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """
        批量执行步骤序列。
        arguments 用于在运行时替换 value_from / tcode_from / system_name_from 参数。
        """
        if arguments is None:
            arguments = {}
        results: Dict[str, Any] = {"success": True, "messages": [], "outputs": {}}

        for step in steps:
            action = step.get("action")
            raw_params = dict(step.get("params", {}))

            # 将 value_from / tcode_from / system_name_from 替换为实际值
            params: Dict[str, Any] = {}
            for key, value in raw_params.items():
                if key == "value_from":
                    params["value"] = arguments.get(value, "")
                elif key == "tcode_from":
                    params["tcode"] = arguments.get(value, "")
                elif key == "system_name_from":
                    params["system_name"] = arguments.get(value, "")
                else:
                    params[key] = value

            success = False
            if action == "open_connection":
                success = self.open_connection(params.get("system_name", ""))
            elif action == "open_tcode":
                success = self.open_tcode(params.get("tcode", ""))
            elif action == "fill_field":
                success = self.set_text(params.get("field", ""), params.get("value", ""))
            elif action == "press_key":
                success = self.press_key(params.get("key", "ENTER"))
            elif action == "click_button":
                success = self.click_button(params.get("button_id", ""))
            elif action == "read_table":
                data = self.get_table_data(params.get("table_id", ""))
                results["outputs"][params.get("output_name", "table_data")] = data
                success = True
            elif action == "get_status":
                results["outputs"][params.get("output_name", "status")] = self.get_status()
                success = True
            elif action == "wait":
                time.sleep(params.get("seconds", 1))
                success = True
            else:
                results["messages"].append(f"未知 action：{action}")

            if not success and params.get("critical", False):
                results["success"] = False
                results["messages"].append(f"关键步骤失败：{action}")
                break

        return results
