# Catapult Window Commander Plugin (v3 - Corrected D-Bus and API)
#
# Description: Lists and manages open windows using the "Window Commander"
#              GNOME Shell extension.
#
# This version corrects D-Bus communication based on the official extension
# documentation and fixes Catapult API usage errors.
#
# It requires the "Window Commander" GNOME Shell extension to be installed
# and enabled. You can find it here:
# https://extensions.gnome.org/extension/4933/window-commander/

import json
import logging
from typing import Any, Dict, Generator, List, Optional

# GTK and D-Bus related imports
from gi.repository import Gio, GLib

# Catapult plugin API imports
from catapult.api import Plugin, SearchResult

# Set up a logger for the plugin
logger = logging.getLogger(__name__)

# Constants for the D-Bus interface, corrected based on documentation
DBUS_DESTINATION = "org.gnome.Shell"
DBUS_INTERFACE_NAME = "org.gnome.Shell.Extensions.WindowCommander"
DBUS_OBJECT_PATH = "/org/gnome/Shell/Extensions/WindowCommander"


class WindowCommanderDBus:
    """A helper class to manage D-Bus communication with the GNOME extension."""

    def __init__(self):
        """Initializes the D-Bus proxy."""
        self._proxy = None
        try:
            # The destination is org.gnome.Shell, not the interface name
            self._proxy = Gio.DBusProxy.new_for_bus_sync(
                Gio.BusType.SESSION,
                Gio.DBusProxyFlags.NONE,
                None,
                DBUS_DESTINATION,
                DBUS_OBJECT_PATH,
                DBUS_INTERFACE_NAME,
                None,
            )
            logger.info("Successfully connected to Window Commander D-Bus service.")
        except GLib.Error as e:
            logger.error(
                "Could not connect to Window Commander D-Bus service. "
                "Please ensure the GNOME Shell extension is installed and enabled. "
                f"Error: {e}"
            )

    def _call_method(self, method_name: str, params: Optional[tuple] = None) -> Any:
        """Generic method to call a D-Bus function."""
        if not self._proxy:
            logger.warning("D-Bus proxy is not available. Cannot call method.")
            return None
        try:
            # The full interface name must be prepended to the method for the call
            full_method = f"{DBUS_INTERFACE_NAME}.{method_name}"
            variant = self._proxy.call_sync(
                full_method,
                GLib.Variant.new_tuple(*(params or ())),
                Gio.DBusCallFlags.NONE,
                -1,
                None,
            )
            if variant:
                # The result is a tuple, we want the first element which is the JSON string
                return variant.unpack()[0]
        except GLib.Error as e:
            logger.error(f"Error calling D-Bus method '{method_name}': {e}")
        return None

    def get_all_windows_with_details(self) -> Optional[List[Dict[str, Any]]]:
        """
        Fetches window list and then gets details for each window individually,
        as required by the new extension API.
        """
        # Step 1: Get the list of basic window info (contains IDs)
        list_json = self._call_method("List")
        if not list_json:
            return None

        try:
            basic_windows = json.loads(list_json)
        except json.JSONDecodeError:
            logger.error("Failed to decode JSON response from List method.")
            return None

        detailed_windows = []
        # Step 2: For each window, get its full details
        for win in basic_windows:
            win_id = win.get("id")
            if not win_id:
                continue

            details_json = self._call_method("GetDetails", (GLib.Variant("u", win_id),))
            if details_json:
                try:
                    details = json.loads(details_json)
                    detailed_windows.append(details)
                except json.JSONDecodeError:
                    logger.error(f"Failed to decode JSON for window details (ID: {win_id}).")
        return detailed_windows

    def execute_action(self, action: str, win_id: int) -> None:
        """Executes a specific window action like activate, close, etc."""
        # Mapping simple actions to D-Bus methods and their parameters
        action_map = {
            "activate": ("Activate", GLib.Variant("u", win_id)),
            "close": ("Close", GLib.Variant("u", win_id), GLib.Variant("b", False)),
            "maximize": ("Maximize", GLib.Variant("u", win_id)),
            "unmaximize": ("Unmaximize", GLib.Variant("u", win_id)),
            "minimize": ("Minimize", GLib.Variant("u", win_id)),
        }
        if action in action_map:
            method_name, *params = action_map[action]
            self._call_method(method_name, tuple(params))
        else:
            logger.warning(f"Unknown window action requested: {action}")


class WindowCommander(Plugin):
    """Catapult plugin for interacting with the Window Commander GNOME extension."""

    title = "Window Commander"
    description = "List, focus, and manage open windows."
    keywords = ["w", "win", "window"]

    def __init__(self):
        super().__init__()
        self.dbus_client = WindowCommanderDBus()

    def search(self, query: str) -> Generator[SearchResult, None, None]:
        """Catapult search handler."""
        trigger_word = ""
        for keyword in self.keywords:
            if query.lower().startswith(keyword + " "):
                trigger_word = keyword + " "
                break
        if not trigger_word:
            return

        search_term = query[len(trigger_word) :].strip().lower()

        windows = self.dbus_client.get_all_windows_with_details()
        if windows is None: # Check for None specifically, as an empty list is valid
            yield SearchResult(
                id="error:no-connection",
                title="Window Commander Not Found or Failed",
                description="Please ensure the GNOME extension is enabled.",
                icon="dialog-warning",
                plugin=self,
                score=100,
                fuzzy=False,
                offset=0,
            )
            return

        for win in windows:
            title = win.get("title", "Untitled Window")
            wm_class = win.get("wm_class", "unknown")
            title_lower = title.lower()
            class_lower = wm_class.lower()

            offset = title_lower.find(search_term)
            if offset == -1:
                offset = class_lower.find(search_term)
                if offset == -1:
                    continue # Skip if search term not found in title or class

            win_id = win.get("id")
            if not win_id:
                continue

            # --- Yield a SearchResult for each available action ---

            # 1. Activate Window
            yield SearchResult(
                id=f"activate:{win_id}",
                title=title,
                description=f"Activate | Class: {wm_class}",
                icon=wm_class,
                plugin=self,
                score=100,
                fuzzy=False,
                offset=offset,
            )

            # 2. Close Window
            yield SearchResult(
                id=f"close:{win_id}",
                title=f"Close: {title}",
                description=f"Close Window | Class: {wm_class}",
                icon="window-close",
                plugin=self,
                score=90,
                fuzzy=False,
                offset=offset,
            )

            # 3. Maximize/Unmaximize
            if win.get("maximized", 0) > 0:
                action_id, action_name, icon_name = "unmaximize", "Unmaximize", "view-restore"
            else:
                action_id, action_name, icon_name = "maximize", "Maximize", "view-fullscreen"

            yield SearchResult(
                id=f"{action_id}:{win_id}",
                title=f"{action_name}: {title}",
                description=f"{action_name} Window | Class: {wm_class}",
                icon=icon_name,
                plugin=self,
                score=80,
                fuzzy=False,
                offset=offset,
            )

    def launch(self, window: Any, id: str) -> None:
        """Called by Catapult when the user selects a result."""
        if id.startswith("error:"):
            return

        try:
            action, win_id_str = id.split(":")
            win_id = int(win_id_str)
            self.dbus_client.execute_action(action, win_id)
        except (ValueError, IndexError) as e:
            logger.error(f"Failed to parse result ID '{id}': {e}")

