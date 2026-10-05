// Shared OpenAI chat call used by the job analysis and resume writing steps.
// Server-only: imported from createServerFn handlers.

// Default model. Override in Vercel with the OPENAI_MODEL environment variable.
export const DEFAULT_MODEL = "gpt-4.1-mini";

export interface ChatOptions {
  apiKey: string;
  system: string;
  user: string;
  maxTokens?: number;
  /** Sampling temperature for non-reasoning models. */
  temperature?: number;
  /** Extra penalties used by the long writing call. */
  penalties?: boolean;
}

/** Sends one JSON-mode chat request and returns the raw message content. */
export async function chatJson(opts: ChatOptions): Promise<string> {
  const model = process.env.OPENAI_MODEL || DEFAULT_MODEL;
  const reasoning = /^(o\d|gpt-5)/i.test(model);
  const body: Record<string, unknown> = {
    model,
    messages: [
      { role: "system", content: opts.system },
      { role: "user", content: opts.user },
    ],
    response_format: { type: "json_object" },
  };
  const maxTokens = opts.maxTokens ?? 32000;
  if (reasoning) {
    body.max_completion_tokens = maxTokens;
  } else {
    body.max_tokens = maxTokens;
    body.temperature = opts.temperature ?? 0.7;
    if (opts.penalties) {
      Object.assign(body, { top_p: 0.95, frequency_penalty: 0.4, presence_penalty: 0.3 });
    }
  }

  // "priority" makes OpenAI respond noticeably faster (at a higher price).
  // Set OPENAI_SERVICE_TIER=default in Vercel to turn it off.
  const tier = process.env.OPENAI_SERVICE_TIER ?? "priority";
  if (tier && tier !== "default") body.service_tier = tier;

  const send = () =>
    fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${opts.apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

  let res = await send();
  if (!res.ok && body.service_tier && res.status === 400) {
    const text = await res.clone().text();
    if (/service_tier/i.test(text)) {
      delete body.service_tier; // model/account doesn't support it — retry at normal speed
      res = await send();
    }
  }

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`OpenAI error ${res.status}: ${text}`);
  }
  const data = await res.json();
  return data?.choices?.[0]?.message?.content ?? "";
}
