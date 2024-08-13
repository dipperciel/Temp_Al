from typing import TypedDict, Literal, Dict, Any


class AlNode(TypedDict):
    node_number: int
    text: str
    type: Literal["NarrativeText"]
    metadata: Dict[str, Any]

