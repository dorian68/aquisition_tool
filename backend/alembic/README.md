Alembic is intentionally left minimal for the MVP.

The API creates tables from SQLAlchemy models at startup via `Base.metadata.create_all`.
For production, initialize Alembic against the models in `app/models/` and replace
startup `create_all` with migrations.

