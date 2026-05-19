from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.core.security import create_access_token
from app.models.database import get_db
from app.models.user import Lead, User
from app.schemas.auth import AuthResponse, GoogleAuthRequest
from app.services.auth_google import GoogleAuthError, GoogleTokenVerifier
from app.services.event_logger import log_event
from app.services.n8n_client import N8nClient

router = APIRouter(prefix="/auth", tags=["auth"])


@router.post("/google", response_model=AuthResponse)
def google_auth(
    request: GoogleAuthRequest,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
) -> AuthResponse:
    try:
        google_user = GoogleTokenVerifier().verify(request.id_token)
    except GoogleAuthError as exc:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail=str(exc)) from exc

    user = db.execute(select(User).where(User.email == google_user.email)).scalars().first()
    if not user:
        user = User(
            email=google_user.email,
            name=google_user.name,
            provider="google",
            provider_user_id=google_user.provider_user_id,
        )
        db.add(user)
        db.flush()
    else:
        user.name = google_user.name or user.name
        user.provider_user_id = google_user.provider_user_id or user.provider_user_id
        db.add(user)

    lead = Lead(
        user_id=user.id,
        email=user.email,
        name=user.name,
        source="csv_dashboard_generator",
        utm_source=request.utm_source,
        utm_medium=request.utm_medium,
        utm_campaign=request.utm_campaign,
        first_upload_id=request.first_upload_id,
        dashboard_id=request.dashboard_id,
    )
    db.add(lead)
    db.commit()
    db.refresh(user)

    payload = {
        "event": "lead_created",
        "lead": {"email": user.email, "name": user.name, "source": "csv_dashboard_generator"},
        "dashboard": {"dashboard_id": request.dashboard_id, "dashboard_type": None},
        "utm": {
            "utm_source": request.utm_source,
            "utm_medium": request.utm_medium,
            "utm_campaign": request.utm_campaign,
        },
    }
    log_event(db, "lead_created", payload, user.id)
    background_tasks.add_task(N8nClient().send_event, "lead_created", payload)

    return AuthResponse(
        user={"id": user.id, "email": user.email, "name": user.name},
        access_token=create_access_token(user),
    )

