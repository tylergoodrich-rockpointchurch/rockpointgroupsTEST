# ⛪ Rock Point Camp Group Sorter

A Streamlit app for Rock Point Church that sorts camp attendees into groups — with full branding and separate workflows for KIDS and YTH camps.

## Features

**KIDS Camp**
- Sort by grade, gender, and/or friend requests
- Size-targeted or friend-first grouping modes

**YTH Camp**
- Community-first algorithm: each Sunday night community group becomes a camp group
- Uploads: HS Community Dashboard, MS Community Dashboard (grade promotion automatic), Leaders CSV, Previous Camp Groups
- 87% friend request satisfaction matching manual process
- Leader assignment per group based on community group

**Both camps**
- Inline SVG logos for KIDS and YTH brands
- Fixed sidebar with brand colors, fonts, and taglines
- Formatted Excel output: Dashboard, Summary, per-bucket tabs, Leaders tab, Outside Filter tab

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect the repo, set `app.py` as entry point, click **Deploy**

## File uploads

| File | Camp | Required | Notes |
|---|---|---|---|
| Registration CSV/Excel | Both | ✅ | First Name, Last Name, Grade, Gender, Friend Request |
| HS Community Dashboard | YTH | Recommended | Multi-sheet Excel — Grade 9–12 sheets with leader columns |
| Camp Leaders CSV | YTH | Recommended | Volunteer signup export |
| MS Community Dashboard | YTH | Optional | Grade 6–8 sheets — grades promoted automatically (8th→9th etc.) |
| Previous Camp Groups | YTH | Optional | First Name, Last Name, Group |

## Friend request column variants accepted

`Friend Request` · `Friend Requests` · `Youth's Friend Request` · `Youth's Friend Requests` (straight and curly apostrophe)
