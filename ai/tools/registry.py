"""
ai.tools.registry
===================
Discovers every :class:`~ai.tools.base.BaseTool` subclass under
``ai.tools`` automatically, and dispatches tool-use requests by name.

This is the concrete mechanism behind "future tools can be added
without modifying the planner or existing tools": adding a tool means
adding one new module under ``ai.tools`` containing one ``BaseTool``
subclass. Nothing in this file needs to change, and nothing in
``ai.workflow`` needs to change either -- both discover capability
through this registry, never through a hardcoded tool list.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 8.2.
"""

from __future__ import annotations

import importlib
import inspect
import logging
import pkgutil
from typing import Dict, List, Optional

from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolError, ToolResult

logger = logging.getLogger(__name__)


class UnknownToolError(Exception):
    """Raised when a dispatch is requested for a tool name with no
    registered implementation."""

    def __init__(self, tool_name: str, known_tools: List[str]) -> None:
        super().__init__(f"Unknown tool '{tool_name}'. Registered tools: {sorted(known_tools)}.")
        self.tool_name = tool_name


class ToolRegistry:
    """Holds every discovered :class:`BaseTool` instance, keyed by name."""

    def __init__(self, tools: Optional[Dict[str, BaseTool]] = None) -> None:
        self._tools: Dict[str, BaseTool] = dict(tools) if tools is not None else {}

    def register(self, tool: BaseTool) -> None:
        """Register one tool instance directly (used by
        :func:`discover_tools` and by tests that need a controlled,
        minimal registry rather than the full auto-discovered set)."""
        self._tools[tool.name] = tool

    def get(self, tool_name: str) -> BaseTool:
        """Look up a tool by name.

        Raises:
            UnknownToolError: If no tool with this name is registered.
        """
        try:
            return self._tools[tool_name]
        except KeyError as exc:
            raise UnknownToolError(tool_name, list(self._tools.keys())) from exc

    def all_tools(self) -> List[BaseTool]:
        """Return every registered tool."""
        return list(self._tools.values())

    def tools_in_category(self, category: ToolCategory) -> List[BaseTool]:
        """Return every registered tool in the given category."""
        return [t for t in self._tools.values() if t.category == category]

    def schemas_for_bedrock(self) -> List[dict]:
        """Return every registered tool's schema in the shape
        :class:`ai.llm.provider.ToolSchema` expects, ready to hand to
        an :class:`~ai.llm.provider.LLMProvider` call."""
        return [
            {"name": t.name, "description": t.description, "input_schema": t.schema}
            for t in self._tools.values()
        ]

    def dispatch(self, tool_name: str, arguments: dict, context: BusinessContext) -> ToolResult:
        """Look up and execute a tool by name.

        Args:
            tool_name: The requested tool's name.
            arguments: Arguments to pass to the tool.
            context: The business data the tool should read from.

        Returns:
            The tool's result on success, or a :class:`ToolResult`
            whose ``summary`` plainly states the tool could not be
            fulfilled (never raises out of this method for a tool-level
            failure -- see the class docstring on why fabricating a
            result is worse than reporting the failure).

        Raises:
            UnknownToolError: If ``tool_name`` has no registered
                implementation. This is a caller/model error (asking
                for a tool that doesn't exist), not a data error, and
                is allowed to propagate so the caller can construct an
                appropriate error message back to the model.
        """
        tool = self.get(tool_name)
        try:
            logger.info("Dispatching tool '%s' with arguments: %s", tool_name, arguments)
            return tool.run(arguments, context)
        except ToolError as exc:
            logger.warning("Tool '%s' could not fulfill the request: %s", tool_name, exc.message)
            return ToolResult(summary=f"The {tool.display_name} could not complete this request: {exc.message}")


def discover_tools(package_name: str = "ai.tools") -> ToolRegistry:
    """Scan ``package_name`` for every :class:`BaseTool` subclass and
    return a :class:`ToolRegistry` containing one instance of each.

    Args:
        package_name: The package to scan. Defaults to ``ai.tools``;
            exposed as a parameter specifically so a test can point
            this at an isolated temp package to prove new tools are
            picked up with zero registry code changes.

    Returns:
        A populated :class:`ToolRegistry`.
    """
    package = importlib.import_module(package_name)
    registry = ToolRegistry()

    for _, module_name, _ in pkgutil.iter_modules(package.__path__):
        if module_name in ("base", "registry"):
            continue  # not tool implementation modules
        full_module_name = f"{package_name}.{module_name}"
        module = importlib.import_module(full_module_name)
        for _, candidate in inspect.getmembers(module, inspect.isclass):
            if issubclass(candidate, BaseTool) and candidate is not BaseTool and candidate.__module__ == full_module_name:
                registry.register(candidate())
                logger.debug("Discovered tool '%s' (%s) in %s", candidate.name, candidate.category.value, full_module_name)

    logger.info("Tool discovery complete: %d tool(s) registered: %s", len(registry.all_tools()), sorted(t.name for t in registry.all_tools()))
    return registry
