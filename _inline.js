
    function init() {
      google.script.run.withSuccessHandler(onUser).withFailureHandler(onError).getCurrentUser();
    }

    function onUser(user) {
      document.getElementById('userInfo').innerText =
        user.name + ' さん（' + user.building + '）' + (user.isAdmin ? ' [管理者]' : '');
      if (user.isAdmin) {
        document.getElementById('adminSection').style.display = 'block';
        loadAdmin();
      }
      loadUnread();
      loadHistory();
    }

    function onError(err) {
      document.getElementById('userInfo').innerText = 'エラー: ' + err.message;
      document.getElementById('unreadList').innerText = '';
      document.getElementById('historyList').innerText = '';
    }

    function badgeClass(importance) {
      return importance === '高' ? 'high' : importance === '中' ? 'mid' : 'low';
    }

    function escapeHtml(str) {
      return String(str).replace(/[&<>"']/g, function (s) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[s];
      });
    }

    function loadUnread() {
      google.script.run.withSuccessHandler(renderUnread).withFailureHandler(onError).getUnreadHandovers();
    }

    function renderUnread(list) {
      const el = document.getElementById('unreadList');
      if (!list || list.length === 0) {
        el.className = 'empty';
        el.innerHTML = '未読の引き継ぎはありません。';
        return;
      }
      el.className = '';
      el.innerHTML = list.map(function (item) {
        return (
          '<div class="card unread">' +
          '<div class="meta">' + item['登録日時'] + ' / ' + item['記入者名'] +
          '<span class="badge ' + badgeClass(item['重要度']) + '">' + item['重要度'] + '</span></div>' +
          '<div class="content">' + escapeHtml(item['内容']) + '</div>' +
          '<button onclick="markRead(\'' + item.ID + '\')">確認しました</button>' +
          '</div>'
        );
      }).join('');
    }

    function markRead(id) {
      google.script.run.withSuccessHandler(function () {
        loadUnread();
        loadHistory();
      }).withFailureHandler(onError).markAsRead(id);
    }

    function submitHandover() {
      const content = document.getElementById('content').value.trim();
      const importance = document.getElementById('importance').value;
      if (!content) { alert('内容を入力してください'); return; }
      google.script.run.withSuccessHandler(function () {
        document.getElementById('content').value = '';
        loadHistory();
        loadUnread();
      }).withFailureHandler(onError).addHandover(content, importance);
    }

    function loadHistory() {
      google.script.run.withSuccessHandler(renderHistory).withFailureHandler(onError).getRecentHandovers();
    }

    function renderHistory(list) {
      const el = document.getElementById('historyList');
      if (!list || list.length === 0) {
        el.className = 'empty';
        el.innerHTML = '記録はまだありません。';
        return;
      }
      el.className = '';
      el.innerHTML = list.map(function (item) {
        return (
          '<div class="card' + (item.isRead ? ' read' : '') + '">' +
          '<div class="meta">' + item['登録日時'] + ' / ' + item['記入者名'] +
          '<span class="badge ' + badgeClass(item['重要度']) + '">' + item['重要度'] + '</span>' +
          (item.isRead ? ' ✅既読' : '') + '</div>' +
          '<div class="content">' + escapeHtml(item['内容']) + '</div>' +
          '</div>'
        );
      }).join('');
    }

    function loadAdmin() {
      google.script.run.withSuccessHandler(renderAdmin).withFailureHandler(onError).getAdminOverview();
    }

    function renderAdmin(data) {
      const el = document.getElementById('adminList');
      if (!data || data.length === 0) {
        el.className = 'empty';
        el.innerHTML = '建物マスタにデータがありません。';
        return;
      }
      el.className = '';
      el.innerHTML = data.map(function (b) {
        const entriesHtml = b.entries.length === 0
          ? '<p class="empty">直近14日の記録はありません</p>'
          : b.entries.map(function (e) {
              return (
                '<div class="card">' +
                '<div class="meta">' + e.date + ' / ' + e.author +
                '<span class="badge ' + badgeClass(e.importance) + '">' + e.importance + '</span></div>' +
                '<div class="content">' + escapeHtml(e.content) + '</div>' +
                '<div class="readstatus">既読: ' + (e.readers.join('、') || 'なし') + '</div>' +
                '<div class="readstatus unread-names">未読: ' + (e.unreaders.join('、') || 'なし') + '</div>' +
                '</div>'
              );
            }).join('');
        return '<h3>' + b.building + '</h3>' + entriesHtml;
      }).join('');
    }

    window.onload = init;
  