@ -0,0 +1,106 @@
{
  "project": {
    "PROJECT_ID": "MDA12345",
    "DEVICE_CODE": "DC123"
  },

  "camera": {
    "AV_LABEL": "f/8",
    "ISO_LABEL": "100",
    "TV_REF_LABEL": "1/60",
    "DELAY_S": 3.0,
    "POST_SHOT_WAIT": 2.5,
    "THUMBNAIL_WIDTH_PX": 180,
    "ZOOM_STEPS": {
      "ZOOM_140_STR": "140",
      "ZOOM_120_STR": "120",
      "ZOOM_110_STR": "110",
      "ZOOM_100_STR": "100",
      "ZOOM_055_STR": "55"
    }
  },

  "excel_header": {
    "Part number": "887402-951 Rev. XXX",
    "Part Description": "Feed, Gateway, Ka-Band",
    "Serial number": "MDA12345",
    "Program type": "OneWeb"
  },

  "categories": {
    "CAT_REF": "reference focus sticker",
    "CAT_A": "septum",
    "CAT_D": "boresight corrugation",
    "CAT_B": "west corrugation aperture",
    "CAT_C": "west corrugation deep",
    "CAT_E": "east corrugation aperture",
    "CAT_F": "east corrugation deep",
    "CAT_G": "north corrugation aperture",
    "CAT_H": "north corrugation deep",
    "CAT_I": "south corrugation aperture",
    "CAT_J": "south corrugation deep",
    "CAT_LBL": "label view",
    "CAT_RXRH": "Rx-RH",
    "CAT_RXRH_WG": "Rx RH WG",
    "CAT_RXLH": "Rx-LH",
    "CAT_RXLH_WG": "Rx LH WG",
    "CAT_BOT_VENT": "Bottom & Vent",
    "CAT_BOTVENT": "Bottom Vent",
    "CAT_TTX_TOP": "Tx-WG Upper",
    "CAT_TTX_BOTTOM": "Tx-WG Bottom",
    "CAT_UPPER_BODY": "Upper Body"
  },

  "orders": {
    "FEATURE_ORDER": [
      "CAT_A",
      "CAT_C", "CAT_E",
      "CAT_H", "CAT_J",
      "CAT_B", "CAT_F", "CAT_G", "CAT_I",
      "CAT_D",
      "CAT_LBL",
      "CAT_TTX_TOP", "CAT_TTX_BOTTOM",
      "CAT_BOTVENT",
      "CAT_RXRH_WG", "CAT_RXLH_WG",
      "CAT_RXRH", "CAT_RXLH",
      "CAT_BOT_VENT",
      "CAT_UPPER_BODY"
    ],

    "ORDER_Z140": [
      "CAT_REF", "CAT_LBL",
      "CAT_RXRH_WG", "CAT_RXLH_WG",
      "CAT_A", "CAT_B", "CAT_C", "CAT_E", "CAT_F", "CAT_H", "CAT_J",
      "CAT_TTX_TOP", "CAT_TTX_BOTTOM",
      "CAT_BOTVENT"
    ],

    "ORDER_Z120": ["CAT_D", "CAT_I"],
    "ORDER_Z055": ["CAT_BOT_VENT", "CAT_RXRH", "CAT_RXLH", "CAT_UPPER_BODY"],
    "ORDER_Z110": ["CAT_G"]
  },

  "tv_map": {
    "CAT_REF": "1/60",
    "CAT_LBL": "1/160",
    "CAT_RXRH": "1/60",
    "CAT_RXRH_WG": "1/60",
    "CAT_RXLH": "1/60",
    "CAT_RXLH_WG": "1/60",
    "CAT_BOT_VENT": "1/160",
    "CAT_BOTVENT": "1/60",
    "CAT_A": "1/40",
    "CAT_D": "1/125",
    "CAT_B": "1/125",
    "CAT_C": "1/80",
    "CAT_E": "1/125",
    "CAT_F": "1/80",
    "CAT_G": "1/125",
    "CAT_H": "1/80",
    "CAT_I": "1/125",
    "CAT_J": "1/80",
    "CAT_TTX_TOP": "1/320",
    "CAT_TTX_BOTTOM": "1/320",
    "CAT_UPPER_BODY": "1/160"
  }
}
