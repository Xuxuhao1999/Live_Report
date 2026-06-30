"""配置持久化：保存目录、站点数据"""
import os
import json
from tkinter import messagebox


def load_config(app):
    """读取保存的目录配置"""
    if os.path.exists(app.config_file):
        try:
            with open(app.config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            legacy = data.get("save_dir", "")
            return {
                "pre_save_dir": data.get("pre_save_dir", legacy),
                "live_save_dir": data.get("live_save_dir", legacy),
                "svc_save_dir": data.get("svc_save_dir", legacy),
                "weather_save_dir": data.get("weather_save_dir", legacy),
                "risk_save_dir": data.get("risk_save_dir", legacy),
                "risk_issue_number": data.get("risk_issue_number", 1),
                "lzy_save_dir": data.get("lzy_save_dir", ""),
            }
        except Exception:
            pass
    return {"pre_save_dir": "", "live_save_dir": "", "svc_save_dir": "",
            "weather_save_dir": "", "risk_save_dir": "", "risk_issue_number": 1,
            "lzy_save_dir": ""}


def save_config(app):
    """保存目录配置"""
    try:
        with open(app.config_file, "w", encoding="utf-8") as f:
            json.dump({
                "pre_save_dir": app.pre_save_dir.get(),
                "live_save_dir": app.live_save_dir.get(),
                "svc_save_dir": app.svc_save_dir.get(),
                "weather_save_dir": app.weather_save_dir.get(),
                "risk_save_dir": app.risk_save_dir.get(),
                "risk_issue_number": app.risk_issue_number,
                "lzy_save_dir": app.lzy_save_dir.get(),
            }, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("配置保存失败", f"无法保存目录配置:\n{e}")


def load_stations(app):
    """读取站点列表"""
    if os.path.exists(app.stations_file):
        try:
            with open(app.stations_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list) and len(data) > 0:
                return data
        except Exception:
            pass
    save_stations(app, app.default_stations)
    return app.default_stations.copy()


def save_stations(app, stations_list=None):
    """保存站点列表"""
    if stations_list is None:
        stations_list = app.stations
    try:
        with open(app.stations_file, "w", encoding="utf-8") as f:
            json.dump(stations_list, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("保存站点失败", f"无法写入站点文件:\n{e}")
