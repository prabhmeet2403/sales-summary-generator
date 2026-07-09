"""
ai.tools.client_lookup
========================
Resolves a client (group) name and reports its profile: which
section(s) it appears in, its POC, and its revenue/margin figures.

This is also the tool the workflow graph's Entity Resolution node uses
to confirm a client name mentioned in conversation actually exists in
this generation's data before deeper analysis proceeds (see
``ai.workflow.nodes.entity_resolution``).
"""

from __future__ import annotations

from typing import ClassVar

from ai.context import BusinessContext
from ai.tools.base import BaseTool, ToolCategory, ToolError, ToolResult


class ClientLookupTool(BaseTool):
    """Looks up a client by name and reports its profile."""

    name: ClassVar[str] = "client_lookup"
    display_name: ClassVar[str] = "Client Lookup Tool"
    description: ClassVar[str] = (
        "Looks up a specific client/group by name and reports which section(s) it "
        "belongs to, its POC, and its total revenue and margin. Use this to confirm a "
        "client exists and to get its overall profile before deeper analysis."
    )
    category: ClassVar[ToolCategory] = ToolCategory.ANALYSIS
    schema: ClassVar[dict] = {
        "type": "object",
        "properties": {
            "client": {"type": "string", "description": "The client/group name to look up."},
        },
        "required": ["client"],
    }

    def run(self, arguments: dict, context: BusinessContext) -> ToolResult:
        client_name = arguments.get("client", "")
        if not client_name:
            raise ToolError("A client name is required.")

        matches = context.groups_df[context.groups_df["group"].str.lower() == client_name.lower()]
        if matches.empty:
            known = ", ".join(context.group_names()[:10])
            raise ToolError(
                f"No client named '{client_name}' was found in this Summary. "
                f"Some known clients include: {known}."
            )

        sections = sorted(matches["section_title"].unique())
        total_revenue = float(matches["revenue"].sum())
        total_margin = float(matches["margin"].sum())
        pocs = sorted({p for p in matches["poc"].dropna().unique()})

        summary_lines = [
            f"Client: {matches.iloc[0]['group']}",
            f"Section(s): {', '.join(sections)}",
            f"POC: {', '.join(pocs) if pocs else 'none recorded'}",
            f"Total revenue: ${total_revenue:,.0f}",
            f"Total margin: ${total_margin:,.0f}",
        ]
        return ToolResult(
            summary="\n".join(summary_lines),
            raw={
                "client": matches.iloc[0]["group"],
                "sections": sections,
                "pocs": pocs,
                "total_revenue": total_revenue,
                "total_margin": total_margin,
            },
        )
