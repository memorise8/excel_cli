"""Skill loader — parses SKILL.md files and generates tool definitions dynamically."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class SkillDef:
    """Parsed skill definition from SKILL.md."""
    name: str
    description: str
    metadata: dict = field(default_factory=dict)
    content: str = ""
    path: Path = field(default_factory=lambda: Path("."))

    @property
    def tools(self) -> list[dict]:
        """Extract tool references from metadata."""
        return self.metadata.get("tools", [])


class SkillLoader:
    """Discovers and parses SKILL.md files from the skills/ directory."""

    def __init__(self, skills_dir: Path | None = None):
        self.skills_dir = skills_dir or Path(__file__).parent.parent.parent / "skills"

    def discover_skills(self) -> list[SkillDef]:
        """Find all SKILL.md files in the skills directory."""
        skills = []
        if not self.skills_dir.exists():
            return skills

        for skill_dir in sorted(self.skills_dir.iterdir()):
            if not skill_dir.is_dir():
                continue

            skill_file = skill_dir / "skill.md"
            if not skill_file.exists():
                # Try SKILL.md (uppercase)
                skill_file = skill_dir / "SKILL.md"
                if not skill_file.exists():
                    continue

            skill = self._parse_skill(skill_file)
            if skill:
                skills.append(skill)

        return skills

    def get_skill(self, name: str) -> SkillDef | None:
        """Get a specific skill by name."""
        for skill in self.discover_skills():
            if skill.name == name:
                return skill
        return None

    def get_skill_names(self) -> list[str]:
        """Get list of all skill names."""
        return [s.name for s in self.discover_skills()]

    def get_skill_descriptions(self) -> dict[str, str]:
        """Get map of skill name → description."""
        return {s.name: s.description for s in self.discover_skills()}

    def _parse_skill(self, path: Path) -> SkillDef | None:
        """Parse a SKILL.md file with YAML frontmatter."""
        content = path.read_text(encoding="utf-8")

        # Extract YAML frontmatter
        frontmatter = {}
        body = content

        fm_match = re.match(r"^---\s*\n(.*?)\n---\s*\n(.*)$", content, re.DOTALL)
        if fm_match:
            fm_text = fm_match.group(1)
            body = fm_match.group(2)
            frontmatter = self._parse_yaml_simple(fm_text)

        name = frontmatter.get("name", path.parent.name)
        description = frontmatter.get("description", "")

        if not description:
            # Try to extract from first heading or paragraph
            for line in body.split("\n"):
                line = line.strip()
                if line and not line.startswith("#"):
                    description = line[:200]
                    break

        return SkillDef(
            name=name,
            description=description,
            metadata=frontmatter,
            content=body,
            path=path,
        )

    def _parse_yaml_simple(self, text: str) -> dict:
        """Simple YAML-like parser for frontmatter (no dependency needed)."""
        result = {}
        current_key = None
        current_list = None

        for line in text.split("\n"):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue

            # List item
            if stripped.startswith("- "):
                if current_list is not None:
                    item_text = stripped[2:].strip()
                    # Handle "- name: value" in list
                    if ": " in item_text:
                        k, v = item_text.split(": ", 1)
                        current_list.append({k.strip(): v.strip()})
                    else:
                        current_list.append(item_text)
                continue

            # Key-value
            if ": " in stripped:
                key, value = stripped.split(": ", 1)
                key = key.strip()
                value = value.strip()

                if value == "":
                    # Next lines might be a list
                    current_key = key
                    current_list = []
                    result[key] = current_list
                else:
                    result[key] = value
                    current_key = None
                    current_list = None

        return result
