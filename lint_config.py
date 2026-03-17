"""
lint_config.py — Lint Config v1：外部化规则配置（Profile + Selector）

JSON 配置文件格式：
{
  "schema": "lex_docx.lint.config.v1",
  "project": "135-caohejing",

  "defaults": {
    "author": "JT",
    "expected_header_shading": "BFBFBF",
    "forbidden_draft_patterns": ["AutoDocs", "MinerU"],
    "common_typos": []
  },

  "profiles": {
    "dd_report_draft": {
      "gate": { "fail_on": ["error"], "max_warn": 999 },
      "check_range": [1, 99999],
      "rules": {
        "jt_note_format":      { "enabled": true,  "severity": "error" },
        "tc_author_check":     { "enabled": true,  "severity": "error", "tc_author": "JT" },
        "table_header_format": { "enabled": true,  "severity": "warn",
                                  "expected_header_shading": "BFBFBF" },
        "no_old_project_refs": { "enabled": true,  "severity": "error",
                                  "forbidden": ["新元房产", "融汇嘉智"] }
      }
    },
    "dd_report_delivery": {
      "extends": "dd_report_draft",
      "gate": { "fail_on": ["error", "warn"], "max_warn": 0 }
    }
  },

  "selectors": [
    { "when": { "path_regex": ".*尽调报告.*working\\\\.docx$" }, "profile": "dd_report_draft" },
    { "when": { "path_regex": ".*对外版.*\\\\.docx$" },          "profile": "dd_report_delivery" }
  ]
}

合并优先级（低→高）：
  defaults < profile（含 extends 链）< CLI 参数

向后兼容：不传 lint_cfg 时 lint.check() 走原有逻辑，无 severity / gate。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


# --------------------------------------------------------------------------- #
# 数据结构                                                                      #
# --------------------------------------------------------------------------- #

@dataclass
class RuleConfig:
    enabled: bool = True
    severity: str = "error"        # "error" | "warn" | "info"
    overrides: dict = field(default_factory=dict)   # 传入 lint fn 的额外 cfg 键


@dataclass
class Gate:
    fail_on: list[str] = field(default_factory=lambda: ["error"])
    max_warn: int = 999


@dataclass
class ResolvedProfile:
    name: str
    gate: Gate
    check_range: tuple[int, int] | None
    rules: dict[str, RuleConfig]    # rule_name → RuleConfig（含被 extends 的规则）
    base_config: dict               # defaults 层，供各规则兜底使用


# --------------------------------------------------------------------------- #
# 加载                                                                          #
# --------------------------------------------------------------------------- #

def load_file(path: str | Path) -> dict:
    """加载并解析 lint config JSON 文件。"""
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    schema = raw.get("schema", "")
    if schema and not schema.startswith("lex_docx.lint.config"):
        raise ValueError(f"未知 schema: {schema!r}")
    return raw


# --------------------------------------------------------------------------- #
# Profile 解析                                                                  #
# --------------------------------------------------------------------------- #

def _merge_rules(base: dict, override: dict) -> dict:
    """深合并两层 rules dict（override 覆盖 base）。"""
    merged = {}
    all_keys = set(base) | set(override)
    for k in all_keys:
        b = base.get(k, {})
        o = override.get(k, {})
        merged[k] = {**b, **o}
    return merged


def _resolve_profile_raw(profiles: dict, name: str,
                          visited: set[str] | None = None) -> dict:
    """递归解析 extends 链，返回合并后的原始 profile dict。"""
    if visited is None:
        visited = set()
    if name in visited:
        raise ValueError(f"lint config: profile extends 形成循环引用: {name!r}")
    visited.add(name)

    profile = dict(profiles.get(name) or {})
    parent_name = profile.pop("extends", None)
    if parent_name:
        parent = _resolve_profile_raw(profiles, parent_name, visited)
        # 合并：parent 为 base，当前 profile 覆盖
        merged_rules = _merge_rules(
            parent.get("rules", {}),
            profile.get("rules", {}),
        )
        profile = {**parent, **profile, "rules": merged_rules}

    return profile


def _build_rule_config(rule_raw: dict) -> RuleConfig:
    """从 raw rule dict 构建 RuleConfig，剩余 key 作为 overrides。"""
    enabled  = bool(rule_raw.get("enabled", True))
    severity = str(rule_raw.get("severity", "error")).lower()
    overrides = {k: v for k, v in rule_raw.items()
                 if k not in ("enabled", "severity")}
    return RuleConfig(enabled=enabled, severity=severity, overrides=overrides)


def _select_profile(raw_cfg: dict, doc_path: str | None) -> str | None:
    """按 selectors[].when.path_regex 匹配文件路径，返回命中的 profile 名。"""
    if not doc_path:
        return None
    selectors = raw_cfg.get("selectors", [])
    for sel in selectors:
        when = sel.get("when", {})
        pat = when.get("path_regex")
        if pat and re.search(pat, doc_path):
            return sel.get("profile")
    return None


def resolve(
    raw_cfg: dict,
    profile_name: str | None = None,
    doc_path: str | None = None,
) -> ResolvedProfile:
    """
    解析 lint config，返回 ResolvedProfile。

    优先级（低→高）：defaults < profile（含 extends）< 显式 profile_name
    若 profile_name 为 None，先尝试 selectors 匹配，再取 profiles 中第一个。
    """
    profiles = raw_cfg.get("profiles", {})
    defaults = raw_cfg.get("defaults", {})

    # 确定 profile 名
    if profile_name is None:
        profile_name = _select_profile(raw_cfg, doc_path)
    if profile_name is None and profiles:
        profile_name = next(iter(profiles))   # 取第一个
    if profile_name is None:
        # 空配置：返回一个默认 ResolvedProfile
        return ResolvedProfile(
            name="default",
            gate=Gate(),
            check_range=None,
            rules={},
            base_config=dict(defaults),
        )

    raw_profile = _resolve_profile_raw(profiles, profile_name)

    # Gate
    gate_raw = raw_profile.get("gate", {})
    gate = Gate(
        fail_on=gate_raw.get("fail_on", ["error"]),
        max_warn=gate_raw.get("max_warn", 999),
    )

    # check_range
    cr_raw = raw_profile.get("check_range")
    check_range = tuple(cr_raw) if cr_raw else None

    # Rules
    rules_raw = raw_profile.get("rules", {})
    rules = {name: _build_rule_config(rc) for name, rc in rules_raw.items()}

    # base_config：defaults 合并进去
    base_config = dict(defaults)

    return ResolvedProfile(
        name=profile_name,
        gate=gate,
        check_range=check_range,
        rules=rules,
        base_config=base_config,
    )


# --------------------------------------------------------------------------- #
# Gate 判定                                                                     #
# --------------------------------------------------------------------------- #

def gate_check(results: list, gate: Gate) -> dict:
    """
    根据 gate 标准判断整体是否通过。

    Args:
        results: list of LintResult（需有 .severity 和 .passed 属性）
        gate:    Gate 实例

    Returns:
        {
          "gate": "PASS" | "FAIL",
          "summary": {"error": N, "warn": N, "info": N},
          "fail_reasons": [...]
        }
    """
    summary: dict[str, int] = {"error": 0, "warn": 0, "info": 0}
    for r in results:
        sev = getattr(r, "severity", "error")
        if not r.passed:
            summary[sev] = summary.get(sev, 0) + 1

    fail_reasons = []
    for sev in gate.fail_on:
        count = summary.get(sev, 0)
        if count > 0:
            fail_reasons.append(f"{count} 条 {sev} 级问题")

    warn_count = summary.get("warn", 0)
    if gate.max_warn < 999 and warn_count > gate.max_warn:
        fail_reasons.append(f"warn 数量 {warn_count} 超过上限 {gate.max_warn}")

    return {
        "gate": "FAIL" if fail_reasons else "PASS",
        "summary": summary,
        "fail_reasons": fail_reasons,
    }


# --------------------------------------------------------------------------- #
# 规则级 config 合并辅助                                                         #
# --------------------------------------------------------------------------- #

# 某些规则的 overrides key 需要重映射到 cfg_dict 的嵌套路径
_RULE_OVERRIDE_REMAP: dict[str, dict[str, str]] = {
    # rule_name -> {override_key: cfg_dict_key}
    # flat key 直接用，无需 remap
}

# no_old_project_refs 的 "forbidden" 需写入 entity_names.forbidden
def apply_rule_overrides(cfg_dict: dict, rule_name: str, overrides: dict) -> dict:
    """将 rule-level overrides 合并到 cfg_dict 副本，返回合并后的 dict。"""
    merged = dict(cfg_dict)
    for k, v in overrides.items():
        if rule_name == "no_old_project_refs" and k == "forbidden":
            en = dict(merged.get("entity_names") or {})
            en["forbidden"] = v
            merged["entity_names"] = en
        else:
            merged[k] = v
    return merged
