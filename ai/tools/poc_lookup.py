"""
ai.tools.poc_lookup
=====================
Rolls up every client under a given POC (account owner): how many
clients, and their combined revenue and margin.
"""

from __future__ import annotations

from typing import ClassVar

from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolError, ToolResult


class POCLookupTool(BaseTool):
    """Looks up a POC (account owner) and reports their book of business."""

    name: ClassVar[str] = "poc_lookup"
    display_name: ClassVar[str] = "POC Lookup Tool"
    description: ClassVar[str] = (
        "Looks up a POC (account owner) by name and reports how many clients they own "
        "and the combined revenue and margin across those clients. Use this for "
        "questions like 'what does Vijay manage' or 'total revenue for Neeraj's accounts'."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {
        "type": "object",
        "properties": {
            "poc": {"type": "string", "description": "The POC (account owner) name to look up."},
        },
        "required": ["poc"],
    }

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        poc_name = arguments.get("poc", "")
        if not poc_name:
            raise ToolError("A POC name is required.")

        matches = context.groups_df[
            context.groups_df["poc"].str.lower() == poc_name.lower()
        ] if "poc" in context.groups_df.columns else context.groups_df.iloc[0:0]

        if matches.empty:
            known = ", ".join(context.poc_names())
            raise ToolError(
                f"No POC named '{poc_name}' was found in this Summary. "
                f"Known POCs include: {known}."
            )

        client_count = matches["group"].nunique()
        total_revenue = float(matches["revenue"].sum())
        total_margin = float(matches["margin"].sum())
        clients = sorted(matches["group"].unique())

        summary = (
            f"POC: {matches.iloc[0]['poc']}\n"
            f"Clients managed: {client_count} ({', '.join(clients)})\n"
            f"Combined revenue: ${total_revenue:,.0f}\n"
            f"Combined margin: ${total_margin:,.0f}"
        )
        return ToolResult(
            summary=summary,
            raw={
                "poc": matches.iloc[0]["poc"],
                "client_count": int(client_count),
                "clients": clients,
                "total_revenue": total_revenue,
                "total_margin": total_margin,
            },
        )
