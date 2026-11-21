# Topo badge records in Terrain

Terrain has nowhere to record that an **additional award** (Camper, Walkabout, First Aid, SES and the rest) was awarded or
presented, and no call a leader can reach creates one. The records that exist are written by something outside Terrain's
web app. Topo therefore keeps its own records **inside Terrain**, as comments on one planned event for the group. Every
copy of Topo uses the same event and the same record format, so any leader in a group, on any copy of Topo, finds and
reads the same records.

The constants below are in `Model/BadgeStore/BadgeStoreSpec.cs`; the reader is `Model/BadgeStore/BadgeStoreRecords.cs`.

## The store event

| | |
|---|---|
| Owner | The **group** (a group event, `invitee_type` `group`, `invitee_id` the group id) |
| Title | `Topo records - please do not delete`. The event is found by the title **starting** `Topo records` (ignoring case), so the end of it can be reworded |
| Start and end | `2100-02-01T08:00:00Z` to `09:00:00Z`, status `planned` |
| Description | `Topo records v1. The Topo app ... Please do not edit or delete this event.` |
| Attendance, challenge area | None. A planned event with no attendance gives nobody milestone credit |

One store per group. If more than one matching event is found, Topo reads all of them and merges the comments.

## Finding it

Terrain's calendar read, `GET events /members/{memberId}/events?start_datetime=&end_datetime=`, only returns events from
calendars that are switched on, and a leader's group calendar is usually off. To find the store:

1. `GET events /members/{memberId}/calendars` (read it fresh).
2. If the group's calendar (`own_calendars`, `type` `group`, `id` the group id) is off, `PUT` the same object back with that
   one `selected` set to `true` (204).
3. Read the calendar for `2100-01-31` to `2100-02-03` and keep the events that match the table above.
4. `PUT` the object from step 1 back **exactly as it was**, even if step 3 failed.

The selection belongs to the account, not to the client, so it also changes what the leader sees in Terrain's own web app
for those moments. Do these steps one at a time. The organiser can also find the event without any of this in
`GET events /members/{memberId}/organised-events`.

## Records

Each record is **one comment** on the event (`POST events /events/{eventId}/comments` with `{ "member_id", "text" }`, 201).
The text is one JSON object:

```json
{"v":1,"t":"award","m":"<member id>","a":"<award id>","d":"2026-08-26","p":"2026-09-02"}
{"v":1,"t":"revoke","m":"<member id>","a":"<award id>"}
```

| Field | Meaning |
|---|---|
| `v` | Format version, an integer. This document is version 1 |
| `t` | `award` records an award; `revoke` takes the record back |
| `m` | The Terrain member id (the opaque UUID). **Never a name** |
| `a` | The Terrain additional award id, for example `camper_award_10`, `walkabout_50`, `first_aid` (the list is `GET templates /additional-awards/specifications.json`) |
| `d` | Date awarded, `yyyy-MM-dd` (required for `award`) |
| `p` | Date presented, `yyyy-MM-dd` (optional) |

Rules for reading:

- For one `m` and `a`, the **latest comment wins** (by Terrain's `created_datetime`, then comment id). An `award` sets the
  record; a `revoke` removes it. A later `award` brings it back.
- Only comments written by an **adult leader** are used (the author is `member.id` on the comment). The event is visible to
  everyone in the group, youth and parents included, and anyone who can open it can comment, so the author is checked.
  Comments from anyone else are ignored and counted so a screen can say so.
- A comment that is not valid JSON, is missing a field, or has an unparseable date is ignored and counted.
- A record with a `v` **higher than the reader understands** is left alone and counted. It is not guessed at.
- Comments with `status` `deleted` are ignored. Unknown extra fields are ignored.
- Reading and writing the same record any number of times is safe.

## Writing

Writing is a comment as above, posted as the signed-in leader (`BadgeStoreService`, `Model/BadgeStore/BadgeStoreWrites.cs`).
Topo always shows the leader what will be written first and asks before changing anything in Terrain.

- **Create the store** (once per group, a group leader only): `POST events /groups/{groupId}/events` with an ordinary planned
  group event: the title, dates and description above, the leader as organiser and verifier, `event_type`
  `{ "type": "group", "id": groupId }`, no attendance, status `planned`. Terrain answers 204 and does not say which event it made,
  so Topo finds it again in the leader's organised events (`GET events /members/{memberId}/organised-events`), which shows it
  whichever calendars are on. It is never created if one already exists.
- **Record an award:** post an `award` comment. Topo reads the store first and does not post it again if the store already
  says the same (a request with no presented date never takes away a presented date already recorded). It then reads the
  comments back to confirm it is there.
- **Take a record back:** post a `revoke`, or delete the comment you wrote (`DELETE events /events/{eventId}/comments/{commentId}`).
- **Which event:** if more than one store event exists, writes go to the one with the lowest id, so two leaders never split
  the records; reads merge all of them.
- A write Terrain refuses can come back to a browser as no reply at all (the events host's rejections can carry no CORS header).
  Topo reports it as a failure and leaves the store as it was.

## What was tested in Terrain (2026-11-22, a leader's own record, everything deleted afterwards)

- A group event is accepted (204); all group members are invitees; it is not in the calendar read unless the group calendar
  is on; comments post, list (newest first, with author) and delete; comments of up to 50,000 characters were accepted.
- Not tested: whether a leader can edit or delete another leader's comment.

## Keep in mind

- This is a workaround for something Terrain does not offer. It can stop working if Terrain changes how events or comments
  behave. Keep a copy of what matters outside Terrain.
- Everyone in the group can see the event and its comments, so records hold only opaque ids and dates.
