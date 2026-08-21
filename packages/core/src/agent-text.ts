export function agentRoundSeparator(previous: string, next: string): string {
  if (!previous || !next) return "";
  const trailing = previous.match(/\n*$/)?.[0].length ?? 0;
  const leading = next.match(/^\n*/)?.[0].length ?? 0;
  return "\n".repeat(Math.max(0, 2 - Math.min(2, trailing) - Math.min(2, leading)));
}

export function appendAgentRoundText(previous: string, next: string): string {
  return previous + agentRoundSeparator(previous, next) + next;
}
