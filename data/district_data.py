WES_ROOMS = []

# --- CLASSROOMS ---
for i in range(1, 28):
    if i in [20, 21]:
        continue  # ECC skip

    if i == 23:
        carpet_fraction = 0.5
        drag = 0
    elif 16 <= i <= 27:
        carpet_fraction = 0.91
        drag = 5
    else:
        carpet_fraction = 0.66
        drag = 0

    WES_ROOMS.append({
        "school": "WES",
        "room_name": f"Room {i}",
        "room_type": "Classroom",
        "area_sqft": 868,
        "carpet": {
            "has_carpet": True,
            "carpet_fraction": carpet_fraction,
            "transition_drag_minutes": drag
        }
    })

# K room
WES_ROOMS.append({
    "school": "WES",
    "room_name": "K",
    "room_type": "Classroom",
    "area_sqft": 1178,
    "carpet": {"has_carpet": True, "carpet_fraction": 0.66, "transition_drag_minutes": 0}
})

# --- SUPPORT SPACES ---
WES_ROOMS += [
    {"room_name": "Staff Room", "room_type": "Staff", "school": "WES",
     "carpet": {"has_carpet": True, "carpet_fraction": 0.5, "transition_drag_minutes": 0}},

    {"room_name": "Conference 1", "room_type": "Office", "school": "WES",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Conference 2", "room_type": "Office", "school": "WES",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Main Office", "room_type": "Office", "school": "WES",
     "carpet": {"has_carpet": True, "carpet_fraction": 0.875, "transition_drag_minutes": 0}},

    {"room_name": "New Gym", "room_type": "Gym", "school": "WES",
     "carpet": {"has_carpet": False}},

    {"room_name": "Old Gym", "room_type": "Gym", "school": "WES",
     "carpet": {"has_carpet": False}},

    {"room_name": "Kitchen Old", "room_type": "Kitchen", "school": "WES",
     "carpet": {"has_carpet": False}},

    {"room_name": "Kitchen New", "room_type": "Kitchen", "school": "WES",
     "carpet": {"has_carpet": False}},

    {"room_name": "Hallways", "room_type": "Hallway", "school": "WES",
     "carpet": {"has_carpet": False}}
]


JXW_ROOMS = []

# KA-KD
for k in ["KA", "KB", "KC", "KD"]:
    frac = 0.66 if k in ["KC", "KD"] else 0.875
    JXW_ROOMS.append({
        "school": "JXW",
        "room_name": k,
        "room_type": "Classroom",
        "carpet": {"has_carpet": True, "carpet_fraction": frac, "transition_drag_minutes": 0}
    })

# 1–27
for i in range(1, 28):
    if i in [25, 26]:
        continue

    if i == 24:
        has_carpet = False
        frac = 0
    elif i in [13,14,15,16,23,27]:
        frac = 0.875
        has_carpet = True
    else:
        frac = 0.875
        has_carpet = True

    JXW_ROOMS.append({
        "school": "JXW",
        "room_name": f"Room {i}",
        "room_type": "Classroom",
        "carpet": {"has_carpet": has_carpet, "carpet_fraction": frac, "transition_drag_minutes": 0}
    })

JXW_ROOMS += [
    {"room_name": "Staff Room", "room_type": "Staff", "school": "JXW",
     "carpet": {"has_carpet": True, "carpet_fraction": 0.5, "transition_drag_minutes": 0}},

    {"room_name": "Main Office", "room_type": "Office", "school": "JXW",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Gym", "room_type": "Gym", "school": "JXW",
     "carpet": {"has_carpet": False}},

    {"room_name": "Kitchen", "room_type": "Kitchen", "school": "JXW",
     "carpet": {"has_carpet": False}},

    {"room_name": "Hallways", "room_type": "Hallway", "school": "JXW",
     "carpet": {"has_carpet": False}}
]


RLS_ROOMS = []

# K rooms
for k in ["K1", "K2"]:
    RLS_ROOMS.append({
        "school": "RLS",
        "room_name": k,
        "room_type": "Classroom",
        "carpet": {"has_carpet": True, "carpet_fraction": 0.875, "transition_drag_minutes": 0}
    })

# K3 K4
for k in ["K3", "K4"]:
    RLS_ROOMS.append({
        "school": "RLS",
        "room_name": k,
        "room_type": "Classroom",
        "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}
    })

# Rooms 3–25
for i in range(3, 26):
    if 3 <= i <= 7:
        frac = 0.875
    elif 8 <= i <= 13:
        frac = 1.0
    else:
        frac = 0.66

    RLS_ROOMS.append({
        "school": "RLS",
        "room_name": f"Room {i}",
        "room_type": "Classroom",
        "carpet": {"has_carpet": True, "carpet_fraction": frac, "transition_drag_minutes": 0}
    })

RLS_ROOMS += [
    {"room_name": "Reading Room", "room_type": "Classroom", "school": "RLS",
     "carpet": {"has_carpet": True, "carpet_fraction": 0.875, "transition_drag_minutes": 0}},

    {"room_name": "Staff Room", "room_type": "Staff", "school": "RLS",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Conference 1", "room_type": "Office", "school": "RLS",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Conference 2", "room_type": "Office", "school": "RLS",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Main Office", "room_type": "Office", "school": "RLS",
     "carpet": {"has_carpet": True, "carpet_fraction": 1.0, "transition_drag_minutes": 0}},

    {"room_name": "Gym", "room_type": "Gym", "school": "RLS",
     "carpet": {"has_carpet": False}},

    {"room_name": "Kitchen", "room_type": "Kitchen", "school": "RLS",
     "carpet": {"has_carpet": False}},

    {"room_name": "Hallways", "room_type": "Hallway", "school": "RLS",
     "carpet": {"has_carpet": False}}
]


ALL_ROOMS = WES_ROOMS + JXW_ROOMS + RLS_ROOMS