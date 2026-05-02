# PACT Mail Queue Flow

This app uses `PACT Mail History` as the outbound queue and audit log.

## Queue contract

Every outbound notification should create a list item with:

- `Title` - email subject
- `RecipientEmail` - one or more recipients, comma-separated
- `MailBody` - full HTML body
- `Status` - `Pending`, `Processing`, `Sent`, or `Failed`

## Processing order

Messages should be processed sequentially in `Created` order, oldest first.

Recommended Flow settings:

- Trigger: recurring schedule, every 1 minute
- Concurrency control: on, degree of parallelism `1`
- Query: `Status eq 'Pending'`
- Sort: `Created asc`

## Flow steps

1. Get the oldest pending item from `PACT Mail History`.
2. Update that item to `Processing`.
3. Send the email to `RecipientEmail`.
4. If send succeeds, update the item to `Sent`.
5. If send fails, update the item to `Failed` and capture the error in the run history.

## App behavior

The SPFx app should only enqueue mail.

That means:

- incident submission creates the SharePoint case record
- case submission also creates a `Pending` mail history entry
- the flow sends mail later
- the dashboard and mail log page read the queue status from SharePoint

## Escalation behavior

The existing case workflow remains unchanged:

- Admin logs the incident
- Repeat-offence tracker updates
- Tier escalation is calculated
- Offender and supervisor are notified through the queue
- Paid status is updated later by admin

## Notes

- Do not send email directly from the browser client for production.
- Use the queue as the system of record.
- Keep list writes idempotent where possible so retries do not duplicate notifications.
