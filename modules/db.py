import os, json
from dotenv import load_dotenv

load_dotenv()

def _get_secret(key):
    try:
        import streamlit as st
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, "")

try:
    from supabase import create_client, Client
    _url = _get_secret("SUPABASE_URL")
    _key = _get_secret("SUPABASE_ANON_KEY")
    supabase: Client = create_client(_url, _key) if _url and _key else None
except Exception:
    supabase = None


def _available():
    return supabase is not None


def save_project(client_name, city, pincode, tier, boq_rows, boq_sources, vendors, notes):
    """Save a generated project to Supabase. Returns project_id or None on failure."""
    if not _available():
        return None
    try:
        proj = supabase.table("projects").insert({
            "client_name": client_name,
            "city": city,
            "pincode": pincode,
            "tier": tier,
        }).execute()
        project_id = proj.data[0]["id"]

        if boq_rows:
            boq_payload = [
                {"project_id": project_id, "category": r.category, "description": r.description,
                 "unit": r.unit, "qty": r.qty, "rate": r.rate}
                for r in boq_rows
            ]
            supabase.table("boq_rows").insert(boq_payload).execute()

        if boq_sources:
            src_payload = [
                {"project_id": project_id, "item": s.item, "basis": s.basis, "source": s.source}
                for s in boq_sources
            ]
            supabase.table("rate_sources").insert(src_payload).execute()

        if vendors:
            vendor_payload = [
                {"project_id": project_id, "category": v.category, "vendor": v.vendor,
                 "specialty": v.specialty, "area": v.area,
                 "lat": v.lat, "lng": v.lng, "rating": v.rating, "phone": v.phone}
                for v in vendors
            ]
            supabase.table("vendors").insert(vendor_payload).execute()

        if notes:
            supabase.table("vendor_notes").insert(
                {"project_id": project_id, "notes": notes}
            ).execute()

        return project_id
    except Exception as e:
        print(f"DB save error: {e}")
        return None


def list_projects():
    """Return list of projects ordered by newest first."""
    if not _available():
        return []
    try:
        res = supabase.table("projects").select("*").order("created_at", desc=True).execute()
        return res.data or []
    except Exception:
        return []


def load_project(project_id):
    """Load full project data by ID. Returns dict with keys: project, boq_rows, rate_sources, vendors, notes."""
    if not _available():
        return None
    try:
        proj = supabase.table("projects").select("*").eq("id", project_id).single().execute()
        boq = supabase.table("boq_rows").select("*").eq("project_id", project_id).execute()
        src = supabase.table("rate_sources").select("*").eq("project_id", project_id).execute()
        vnd = supabase.table("vendors").select("*").eq("project_id", project_id).execute()
        notes = supabase.table("vendor_notes").select("notes").eq("project_id", project_id).execute()
        return {
            "project": proj.data,
            "boq_rows": boq.data or [],
            "rate_sources": src.data or [],
            "vendors": vnd.data or [],
            "notes": notes.data[0]["notes"] if notes.data else "",
        }
    except Exception:
        return None


def delete_project(project_id):
    if not _available():
        return False
    try:
        supabase.table("projects").delete().eq("id", project_id).execute()
        return True
    except Exception:
        return False
