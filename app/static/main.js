// Dashboard Chart.js rendering
// Data is injected via window.CHART_DATA from dashboard.html

(function () {
  if (typeof CHART_DATA === 'undefined') return;

  var COLORS = [
    'rgba(59,130,246,0.85)',
    'rgba(34,197,94,0.85)',
    'rgba(245,158,11,0.85)',
    'rgba(239,68,68,0.85)',
    'rgba(168,85,247,0.85)',
    'rgba(6,182,212,0.85)',
  ];

  var CHART_DEFAULTS = {
    plugins: {
      legend: {
        labels: { color: '#e7e7ea', font: { size: 12 } }
      }
    },
    scales: {
      x: {
        grid: { color: '#2b2b36' },
        ticks: { color: '#8888a0' }
      },
      y: {
        grid: { color: '#2b2b36' },
        ticks: { color: '#8888a0' }
      }
    }
  };

  // Doughnut: Models by Tier
  var tierCtx = document.getElementById('tierChart');
  if (tierCtx && CHART_DATA.tier_labels) {
    new Chart(tierCtx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: CHART_DATA.tier_labels,
        datasets: [{
          data: CHART_DATA.tier_values,
          backgroundColor: COLORS,
          borderColor: '#1e1e24',
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: '#e7e7ea', font: { size: 12 }, padding: 16 }
          }
        }
      }
    });
  }

  // Bar: Models by Risk Type
  var riskCtx = document.getElementById('riskChart');
  if (riskCtx && CHART_DATA.risk_labels) {
    new Chart(riskCtx.getContext('2d'), {
      type: 'bar',
      data: {
        labels: CHART_DATA.risk_labels,
        datasets: [{
          label: 'Models',
          data: CHART_DATA.risk_values,
          backgroundColor: COLORS,
          borderRadius: 4,
          borderSkipped: false
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          x: {
            grid: { color: '#2b2b36' },
            ticks: { color: '#8888a0' }
          },
          y: {
            beginAtZero: true,
            grid: { color: '#2b2b36' },
            ticks: { color: '#8888a0', stepSize: 1 }
          }
        }
      }
    });
  }
})();
