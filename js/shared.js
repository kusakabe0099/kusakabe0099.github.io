/* 宛先・添付の収集、確認済みフラグの保存/取得（commands.js と sendCheck.js で共有） */
(function (global) {
  "use strict";

  const CONFIRMED_KEY = "okanConfirmedSignature";

  function call(fn) {
    return new Promise((resolve, reject) => {
      fn((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(result.error);
      });
    });
  }

  // 短い署名文字列を作る（djb2）。宛先・添付が変わると値が変わる。
  function hash(text) {
    let h = 5381;
    for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0;
    return (h >>> 0).toString(16);
  }

  function domainOf(address) {
    const at = address.lastIndexOf("@");
    return at >= 0 ? address.slice(at + 1).toLowerCase() : "(不明)";
  }

  // 送信者と同じドメインなら社内、それ以外は社外。
  // 送信者のドメインが分からないときは、安全側に倒して社外として扱う。
  function isExternal(domain, senderDomain) {
    return !senderDomain || domain !== senderDomain;
  }

  // 送信者（From）のドメインを取得する。
  // 共有メールボックスや「代理で送信」の場合は、From に指定されているアドレスを優先する。
  async function getSenderDomain(item) {
    try {
      const from = await call((cb) => item.from.getAsync(cb));
      if (from && from.emailAddress) return domainOf(from.emailAddress.toLowerCase());
    } catch (e) {
      // 取得できない場合は、下の「サインイン中のユーザー」にフォールバックする
    }

    const profile = Office.context.mailbox && Office.context.mailbox.userProfile;
    if (profile && profile.emailAddress) return domainOf(profile.emailAddress.toLowerCase());

    return null;
  }

  async function collect(item) {
    const [to, cc, bcc, attachments, senderDomain] = await Promise.all([
      call((cb) => item.to.getAsync(cb)),
      call((cb) => item.cc.getAsync(cb)),
      call((cb) => item.bcc.getAsync(cb)),
      call((cb) => item.getAttachmentsAsync(cb)),
      getSenderDomain(item),
    ]);

    const recipients = [];
    [["To", to], ["Cc", cc], ["Bcc", bcc]].forEach(([kind, list]) => {
      list.forEach((r) =>
        recipients.push({
          kind,
          name: r.displayName || "",
          address: (r.emailAddress || "").toLowerCase(),
          external: isExternal(domainOf((r.emailAddress || "").toLowerCase()), senderDomain),
        })
      );
    });

    const files = attachments
      .filter((a) => !a.isInline)
      .map((a) => ({ name: a.name, size: a.size }));

    // ドメインごとにまとめる（社外を先頭に）
    const map = new Map();
    recipients.forEach((r) => {
      const domain = domainOf(r.address);
      if (!map.has(domain)) map.set(domain, { domain, external: isExternal(domain, senderDomain), recipients: [] });
      map.get(domain).recipients.push(r);
    });
    const groups = Array.from(map.values()).sort(
      (a, b) => Number(b.external) - Number(a.external) || a.domain.localeCompare(b.domain)
    );

    const signature = hash(
      JSON.stringify([
        recipients.map((r) => r.kind + ":" + r.address).sort(),
        files.map((f) => f.name + ":" + f.size).sort(),
        senderDomain,
      ])
    );

    return {
      recipients,
      groups,
      files,
      externalCount: recipients.filter((r) => r.external).length,
      senderDomain,
      signature,
    };
  }

  async function getConfirmed(item) {
    try {
      return await call((cb) => item.sessionData.getAsync(CONFIRMED_KEY, cb));
    } catch (e) {
      return null; // 未保存
    }
  }

  function setConfirmed(item, signature) {
    return call((cb) => item.sessionData.setAsync(CONFIRMED_KEY, signature, cb));
  }

  global.Okan = { call, collect, domainOf, isExternal, getConfirmed, setConfirmed };
})(window);
