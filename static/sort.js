// sort.js - 표 정렬 기능
// class="sortable" 이 붙은 표의 머리글(th)을 누르면 그 열 기준으로 정렬합니다.
// 한 번 누르면 오름차순(▲), 다시 누르면 내림차순(▼).
// 정렬하면 안 되는 열(관리 버튼 등)은 th 에 data-nosort 를 붙입니다.

document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('table.sortable').forEach(function (table) {
    var headers = table.querySelectorAll('thead th');

    headers.forEach(function (th, col) {
      if (th.hasAttribute('data-nosort')) return;
      th.classList.add('sort-th');
      th.title = '눌러서 정렬';

      th.addEventListener('click', function () {
        var asc = th.getAttribute('data-dir') !== 'asc';
        headers.forEach(function (h) { h.removeAttribute('data-dir'); });
        th.setAttribute('data-dir', asc ? 'asc' : 'desc');

        var tbody = table.tBodies[0];
        // '내용 없음' 줄(class="empty")은 정렬에서 제외
        var rows = Array.prototype.filter.call(tbody.rows, function (r) {
          return !r.classList.contains('empty');
        });

        rows.sort(function (a, b) {
          var x = a.cells[col].innerText.trim();
          var y = b.cells[col].innerText.trim();
          // numeric: true → '10' 이 '9' 보다 뒤에 오도록 숫자처럼 비교
          var result = x.localeCompare(y, 'ko', { numeric: true });
          return asc ? result : -result;
        });

        rows.forEach(function (r) { tbody.appendChild(r); });
      });
    });
  });
});
