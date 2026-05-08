# jd-bot-service (JD-only extraction)

This is a minimal extraction of the JD flows from the main project. It supports:
- JD Creation (multi-step and single-form variants)
- Fetch JD (department/role filter → list results)

## Run locally

1. Install deps
```
cd jd-bot-service
npm install
```

2. Configure
- `.env` file with Bot Framework and DB settings

3. Start bot
```
npm start
```
The bot listens on `http://localhost:3978`.

4. Expose with ngrok (optional for Teams/Azure Bot Service)
```
ngrok http 3978
```
Set Azure Bot messaging endpoint to `https://<ngrok-subdomain>.ngrok-free.app/api/messages`.

## Notes
- DB schema/queries mirror the main project (`gpm.hrm_modules`, `hrm_job_master`, `pcollab_members`).
- Action IDs and card payloads are unchanged for compatibility.
- No proactive messaging, LLM agents, or meeting orchestration are included.
