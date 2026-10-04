"""VoiceForge v2 end-to-end demo: voice call + live-call vision step.

Stdlib only, zero credentials, zero network. Simulates:
  1. an inbound call with the hardened booking flow (confirm loop),
  2. an escalation -> warm transfer announcement,
  3. the vision flow: photo request -> simulated MMS -> gated analysis ->
     turn-boundary injection,
  4. an outbound campaign with normalized DNC handling.
"""

import datetime as dt
import json

from voiceforge import (
    CallLogger, ConversationBrain, MockLLMClient, TwilioAdapter,
    VapiAdapter, VoiceAgent, VoiceForgeConfig,
)
from voiceforge.providers import MockVisionClient
from voiceforge.vision import (
    ContextInjector, MmsWatch, VisionAnalyzer, VisionFlow,
)
from voiceforge.vision.mms_watch import MmsEvent
from voiceforge.session import MediaItem


def main():
    today = dt.date(2026, 10, 4)
    config = VoiceForgeConfig(business="Swift Towing Co.",
                              transfer_number="+18885551212")
    logger = CallLogger()
    brain = ConversationBrain(business="Swift Towing Co.", config=config,
                              today=today,
                              llm=MockLLMClient({"fallback":
                                                 "Got it. Anything else?"}))
    agent = VoiceAgent(brain, logger=logger, config=config)
    twilio = TwilioAdapter(auth_token="demo-secret",
                           gather_action_url="/voice")

    print("=== 1. INBOUND CALL (hardened booking) ===")
    print(agent.inbound_call("CA_INBOUND1", "+15551234567",
                             twilio=twilio)[:100], "...\n")
    script = [("Hi, what are your hours?", None),
              ("I need a tow booked please", None),
              ("John Carter", None),
              ("flatbed tow", None),
              ("tomorrow", None),
              ("yes", None)]
    for caller, _ in script:
        reply = agent.caller_said("CA_INBOUND1", caller, twilio=twilio)
        print(f"Caller : {caller}\nAgent  : {reply[:110]}\n")

    print("=== 2. ESCALATION -> WARM TRANSFER ===")
    agent.inbound_call("CA_ESC1", "+15557654321")
    out = agent.caller_said("CA_ESC1", "Let me talk to a human agent",
                            twilio=twilio)
    print("Transfer TwiML has <Dial>:", "<Dial" in out)
    print("Whisper context:", agent.build_handoff_context(
        agent.calls["CA_ESC1"])[:100], "...\n")

    print("=== 3. LIVE-CALL VISION STEP ===")
    watch = MmsWatch(twilio, logger, config, media_dir="/tmp/vf2demo",
                     downloader=lambda url: (b"FAKEIMG", "image/jpeg"))
    watch._sleep = lambda s: None
    flow = VisionFlow(
        watch,
        VisionAnalyzer(MockVisionClient(), logger, config), 
        ContextInjector(logger), logger, config)
    agent.vision_flow = flow
    agent.inbound_call("CA_VIS1", "+15550001111")
    ask = agent.request_photo("CA_VIS1", "the cracked windshield")
    print("Agent asks:", ask)
    # simulate the MMS arriving (bypassing HTTP): build the event directly
    import os
    os.makedirs("/tmp/vf2demo", exist_ok=True)
    with open("/tmp/vf2demo/SM9_0.jpg", "wb") as fh:
        fh.write(b"FAKEIMG")
    event = MmsEvent(message_sid="SM9", from_phone="+15550001111",
                     media=[MediaItem(media_sid="SM9_0", message_sid="SM9",
                                      content_type="image/jpeg", size=7,
                                      local_path="/tmp/vf2demo/SM9_0.jpg")])
    watch.seen_sids.add("SM9")
    watch.attach_to_session(agent, event)
    flow._analyze_attached(agent, agent.calls["CA_VIS1"])
    session = agent.calls["CA_VIS1"]
    print("vision_state:", session.vision_state)
    reply = agent.caller_said("CA_VIS1", "Did you get the photo?")
    vision_turns = [t.text for t in session.turns
                    if t.text.startswith("[vision:")]
    print("vision context injected:", len(vision_turns) == 1)
    print("Agent continues:", reply[:90], "\n")

    print("=== 4. VAPI WEBHOOK (fail-closed) ===")
    try:
        VapiAdapter(webhook_secret="")
        print("ERROR: empty secret accepted!")
    except ValueError as exc:
        print("empty secret rejected:", str(exc)[:60], "...")
    vapi = VapiAdapter(webhook_secret="vapi-secret")
    print("verify ok:", vapi.verify({"x-vapi-secret": "vapi-secret"}), "\n")

    print("=== 5. OUTBOUND CAMPAIGN (normalized DNC) ===")
    csv_data = ("phone,name\n+1-555-000-0001,Alice\n"
                "+15550000002,Bob\n+15550000003,Carol\n")
    report = agent.outbound_campaign(
        csv_data,
        opener="Hi {name}, this is the AI assistant for Swift Towing Co.",
        throttle_per_min=600,
        dnc=["+15550000002"],  # matches Bob despite different formatting
        on_transcript=lambda opener: "Yes, what are your prices?",
        sleep=lambda s: None)
    for row in report:
        print({k: row[k] for k in ("phone", "status")})

    print("\n=== DAY SUMMARY ===")
    print(json.dumps(agent.summary(), indent=2))


if __name__ == "__main__":
    main()
