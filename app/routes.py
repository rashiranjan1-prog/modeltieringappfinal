import json
import csv
import io
import os
from datetime import datetime
from flask import (Blueprint, render_template, redirect, url_for, flash,
                   request, abort, Response, current_app)
from .db import get_db
from .auth import login_required, roles_required, has_role, get_current_user
from .services.tieringlogic import compute_tiering_for_model, compute_tiering_for_all

main_bp = Blueprint('main', __name__)

# ── Helpers ───────────────────────────────────────────────────────────────────

GRP_ORDER = "CASE grp WHEN 'Materiality' THEN 1 WHEN 'Criticality' THEN 2 WHEN 'Complexity' THEN 3 ELSE 4 END"


# ── Dashboard ─────────────────────────────────────────────────────────────────

@main_bp.route('/')
@login_required
def dashboard():
    db = get_db()
    models = db.execute('SELECT * FROM models').fetchall()
    tiers  = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()

    tier_counts = {t['name']: 0 for t in tiers}
    for m in models:
        ct = m['current_tier']
        if ct:
            tier_counts[ct] = tier_counts.get(ct, 0) + 1

    risk_counts = {}
    for m in models:
        rt = m['risk_type'] or 'Unknown'
        risk_counts[rt] = risk_counts.get(rt, 0) + 1

    chart_data = {
        'tier_labels': list(tier_counts.keys()),
        'tier_values': list(tier_counts.values()),
        'risk_labels': list(risk_counts.keys()),
        'risk_values': list(risk_counts.values()),
    }
    return render_template('dashboard.html', models=models, tier_counts=tier_counts,
                           chart_data_json=json.dumps(chart_data))


# ── Models ────────────────────────────────────────────────────────────────────

@main_bp.route('/models')
@login_required
def models_list():
    db = get_db()
    models = db.execute('SELECT * FROM models').fetchall()
    return render_template('models.html', models=models)


@main_bp.route('/models/new', methods=['GET', 'POST'])
@login_required
@roles_required('Admin', 'Developer')
def model_new():
    db = get_db()
    parameters = db.execute(f'SELECT * FROM parameters ORDER BY {GRP_ORDER}, id').fetchall()
    groups = {}
    for p in parameters:
        groups.setdefault(p['grp'], []).append(p)

    if request.method == 'POST':
        name      = request.form.get('name', '').strip()
        risk_type = request.form.get('risk_type', '').strip()
        if not name:
            flash('Model name is required.', 'danger')
            return render_template('modelnew.html', groups=groups, parameters=parameters)
        cur = db.execute('INSERT INTO models (name, risk_type) VALUES (?,?)', (name, risk_type))
        model_id = cur.lastrowid
        db.commit()

        # Save parameter weights and scores if provided
        for p in parameters:
            # Update weight if changed
            new_weight = request.form.get(f'weight_{p["id"]}', '').strip()
            if new_weight:
                try:
                    w = float(new_weight)
                    if 0 < w <= 1:
                        db.execute('UPDATE parameters SET weight=? WHERE id=?', (w, p['id']))
                except ValueError:
                    pass

            # Save level score
            val = request.form.get(f'score_{p["id"]}', '').strip()
            if val:
                try:
                    level = max(1, min(3, int(val)))
                    db.execute(
                        'INSERT OR REPLACE INTO model_scores (model_id, parameter_id, level) VALUES (?,?,?)',
                        (model_id, p['id'], level)
                    )
                except ValueError:
                    pass
        db.commit()

        # Manual score override
        manual_score = request.form.get('manual_score', '').strip()
        if manual_score:
            try:
                ms = round(max(1.0, min(3.0, float(manual_score))), 2)
                db.execute('UPDATE models SET computed_score=? WHERE id=?', (ms, model_id))
                db.commit()
            except ValueError:
                pass

        # Compute tiering
        from .services.tieringlogic import compute_tiering_for_model
        compute_tiering_for_model(model_id)

        flash(f'Model "{name}" created.', 'success')
        return redirect(url_for('main.model_detail', model_id=model_id))

    return render_template('modelnew.html', groups=groups, parameters=parameters)


@main_bp.route('/models/<int:model_id>')
@login_required
def model_detail(model_id):
    db = get_db()
    model = db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()
    if not model:
        abort(404)

    parameters  = db.execute(
        f'SELECT * FROM parameters ORDER BY {GRP_ORDER}, id'
    ).fetchall()
    scores_rows = db.execute('SELECT * FROM model_scores WHERE model_id=?', (model_id,)).fetchall()
    scores_map  = {s['parameter_id']: s for s in scores_rows}
    overrides   = db.execute(
        'SELECT o.*, u.email as user_email FROM overrides o JOIN users u ON u.id=o.user_id '
        'WHERE o.model_id=? ORDER BY o.created_at DESC', (model_id,)
    ).fetchall()
    tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()

    groups = {}
    for p in parameters:
        groups.setdefault(p['grp'], []).append(p)

    return render_template('modeldetail.html',
                           model=model, parameters=parameters, groups=groups,
                           scores_map=scores_map, overrides=overrides,
                           tiers=tiers, has_role=has_role)


@main_bp.route('/models/<int:model_id>/scores', methods=['POST'])
@login_required
@roles_required('Admin', 'Developer')
def model_save_scores(model_id):
    db = get_db()
    if not db.execute('SELECT id FROM models WHERE id=?', (model_id,)).fetchone():
        abort(404)

    parameters = db.execute('SELECT * FROM parameters').fetchall()

    # Remove stale rows for deleted parameters
    db.execute(
        'DELETE FROM model_scores WHERE model_id=? '
        'AND parameter_id NOT IN (SELECT id FROM parameters)', (model_id,)
    )

    for p in parameters:
        raw = request.form.get(f'level_{p["id"]}')
        level = max(1, min(3, int(raw))) if raw else 1
        existing = db.execute(
            'SELECT id FROM model_scores WHERE model_id=? AND parameter_id=?',
            (model_id, p['id'])
        ).fetchone()
        if existing:
            db.execute('UPDATE model_scores SET level=? WHERE model_id=? AND parameter_id=?',
                       (level, model_id, p['id']))
        else:
            db.execute('INSERT INTO model_scores (model_id, parameter_id, level) VALUES (?,?,?)',
                       (model_id, p['id'], level))
    db.commit()
    compute_tiering_for_model(model_id)
    flash('Scores saved and tiering computed.', 'success')
    return redirect(url_for('main.model_detail', model_id=model_id))


@main_bp.route('/models/<int:model_id>/delete', methods=['POST'])
@login_required
def model_delete(model_id):
    db = get_db()
    db.execute('DELETE FROM model_scores WHERE model_id=?', (model_id,))
    db.execute('DELETE FROM overrides WHERE model_id=?', (model_id,))
    db.execute('DELETE FROM models WHERE id=?', (model_id,))
    db.commit()
    flash('Model deleted.', 'info')
    return redirect(url_for('main.models_list'))


@main_bp.route('/models/<int:model_id>/override', methods=['POST'])
@login_required
@roles_required('Admin', 'Developer')
def model_override(model_id):
    db = get_db()
    model = db.execute('SELECT * FROM models WHERE id=?', (model_id,)).fetchone()
    if not model:
        abort(404)
    new_tier = request.form.get('new_tier', '').strip()
    reason   = request.form.get('reason', '').strip()
    if not new_tier or not reason:
        flash('Tier and reason are required.', 'danger')
        return redirect(url_for('main.model_detail', model_id=model_id))
    user = get_current_user()
    db.execute('INSERT INTO overrides (model_id, user_id, old_tier, new_tier, reason) VALUES (?,?,?,?,?)',
               (model_id, user.id, model['current_tier'], new_tier, reason))
    db.execute('UPDATE models SET current_tier=? WHERE id=?', (new_tier, model_id))
    db.commit()
    flash(f'Tier overridden to {new_tier}.', 'success')
    return redirect(url_for('main.model_detail', model_id=model_id))


# ── Parameters ────────────────────────────────────────────────────────────────

@main_bp.route('/parameters')
@login_required
def parameters_list():
    db = get_db()
    params = db.execute(
        f'SELECT * FROM parameters ORDER BY {GRP_ORDER}, id'
    ).fetchall()
    groups = {}
    for p in params:
        groups.setdefault(p['grp'], []).append(p)
    return render_template('parameters.html', parameters=params, groups=groups,
                           has_role=has_role)


@main_bp.route('/parameters/new', methods=['GET', 'POST'])
@login_required
@roles_required('Admin', 'Developer')
def parameter_new():
    if request.method == 'POST':
        db = get_db()
        db.execute(
            'INSERT INTO parameters '
            '(grp, sub_parameter, criteria, description, weight,'
            ' level1_label, level2_label, level3_label) VALUES (?,?,?,?,?,?,?,?)',
            (request.form.get('group', '').strip(),
             request.form.get('sub_parameter', '').strip(),
             request.form.get('criteria', '').strip(),
             request.form.get('description', '').strip(),
             float(request.form.get('weight', 1.0)),
             request.form.get('level1_label', '').strip(),
             request.form.get('level2_label', '').strip(),
             request.form.get('level3_label', '').strip())
        )
        db.commit()
        compute_tiering_for_all()
        flash('Parameter created and scores recomputed.', 'success')
        return redirect(url_for('main.parameters_list'))

    db = get_db()
    params = db.execute(f'SELECT * FROM parameters ORDER BY {GRP_ORDER}, id').fetchall()
    groups = {}
    for p in params:
        groups.setdefault(p['grp'], []).append(p)
    return render_template('parameters.html', parameters=params, groups=groups,
                           show_form=True, has_role=has_role)


@main_bp.route('/parameters/<int:param_id>/edit', methods=['GET', 'POST'])
@login_required
@roles_required('Admin', 'Developer')
def parameter_edit(param_id):
    db = get_db()
    param = db.execute('SELECT * FROM parameters WHERE id=?', (param_id,)).fetchone()
    if not param:
        abort(404)
    if request.method == 'POST':
        db.execute(
            'UPDATE parameters SET grp=?, sub_parameter=?, criteria=?, description=?,'
            ' weight=?, level1_label=?, level2_label=?, level3_label=? WHERE id=?',
            (request.form.get('group', '').strip(),
             request.form.get('sub_parameter', '').strip(),
             request.form.get('criteria', '').strip(),
             request.form.get('description', '').strip(),
             float(request.form.get('weight', 1.0)),
             request.form.get('level1_label', '').strip(),
             request.form.get('level2_label', '').strip(),
             request.form.get('level3_label', '').strip(),
             param_id)
        )
        db.commit()
        compute_tiering_for_all()
        flash('Parameter updated and scores recomputed.', 'success')
        return redirect(url_for('main.parameters_list'))

    params = db.execute(f'SELECT * FROM parameters ORDER BY {GRP_ORDER}, id').fetchall()
    groups = {}
    for p in params:
        groups.setdefault(p['grp'], []).append(p)
    return render_template('parameters.html', parameters=params, groups=groups,
                           edit_param=param, has_role=has_role)


@main_bp.route('/parameters/<int:param_id>/delete', methods=['POST'])
@login_required
@roles_required('Admin')
def parameter_delete(param_id):
    db = get_db()
    db.execute('DELETE FROM parameters WHERE id=?', (param_id,))
    db.commit()
    flash('Parameter deleted.', 'success')
    return redirect(url_for('main.parameters_list'))


# ── Tiering ───────────────────────────────────────────────────────────────────

@main_bp.route('/tiering')
@login_required
def model_tiering():
    db = get_db()
    models = db.execute('SELECT * FROM models').fetchall()
    tiers  = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
    score_counts = {
        row['model_id']: row['cnt']
        for row in db.execute(
            'SELECT model_id, COUNT(*) as cnt FROM model_scores GROUP BY model_id'
        ).fetchall()
    }
    return render_template('modeltiering.html', models=models, tiers=tiers,
                           score_counts=score_counts, has_role=has_role)


@main_bp.route('/tiering/run', methods=['POST'])
@login_required
@roles_required('Admin', 'Developer')
def run_tiering():
    count = compute_tiering_for_all()
    flash(f'Tiering computed for {count} models.', 'success')
    return redirect(url_for('main.model_tiering'))


# ── Reports ───────────────────────────────────────────────────────────────────

@main_bp.route('/reports')
@login_required
def reports_list():
    db   = get_db()
    user = get_current_user()
    reports = db.execute(
        'SELECT r.*, u.email as user_email FROM reports r JOIN users u ON u.id=r.user_id '
        'WHERE r.user_id=? ORDER BY r.created_at DESC', (user.id,)
    ).fetchall()
    return render_template('reports.html', reports=reports)


@main_bp.route('/reports/tiering')
@login_required
def report_tiering():
    db          = get_db()
    risk_filter = request.args.get('risk_type', '')
    tier_filter = request.args.get('tier', '')

    query  = 'SELECT * FROM models WHERE 1=1'
    params = []
    if risk_filter:
        query += ' AND risk_type=?'
        params.append(risk_filter)
    if tier_filter:
        query += ' AND current_tier=?'
        params.append(tier_filter)

    models     = db.execute(query, params).fetchall()
    risk_types = [r[0] for r in db.execute(
        'SELECT DISTINCT risk_type FROM models WHERE risk_type IS NOT NULL'
    ).fetchall()]
    tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()

    return render_template('reporttiering.html',
                           models=models, risk_types=risk_types, tiers=tiers,
                           risk_filter=risk_filter, tier_filter=tier_filter,
                           has_role=has_role)


@main_bp.route('/reports/save', methods=['POST'])
@login_required
def report_save():
    db   = get_db()
    user = get_current_user()
    name    = request.form.get('name', '').strip() or f'Report {datetime.utcnow().strftime("%Y-%m-%d")}'
    filters = json.dumps({'risk_type': request.form.get('risk_type', ''),
                          'tier': request.form.get('tier', '')})
    db.execute('INSERT INTO reports (user_id, name, filters_json) VALUES (?,?,?)',
               (user.id, name, filters))
    db.commit()
    flash('Report saved.', 'success')
    return redirect(url_for('main.reports_list'))


@main_bp.route('/reports/export/csv')
@login_required
def report_export_csv():
    db          = get_db()
    risk_filter = request.args.get('risk_type', '')
    tier_filter = request.args.get('tier', '')

    query  = 'SELECT * FROM models WHERE 1=1'
    params = []
    if risk_filter:
        query += ' AND risk_type=?'
        params.append(risk_filter)
    if tier_filter:
        query += ' AND current_tier=?'
        params.append(tier_filter)

    models = db.execute(query, params).fetchall()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Model Name', 'Risk Type', 'Computed Score', 'Computed Tier', 'Current Tier'])
    for m in models:
        writer.writerow([m['name'], m['risk_type'], m['computed_score'],
                         m['computed_tier'], m['current_tier']])

    return Response(output.getvalue(), mimetype='text/csv',
                    headers={'Content-Disposition': 'attachment; filename=tiering_report.csv'})


# ── Upload ────────────────────────────────────────────────────────────────────

@main_bp.route('/upload', methods=['GET', 'POST'])
@login_required
@roles_required('Admin', 'Developer')
def upload():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename.endswith('.xlsx'):
            flash('Please upload a valid .xlsx file.', 'danger')
            return render_template('upload.html')
        upload_dir = current_app.config['UPLOAD_FOLDER']
        filepath   = os.path.join(upload_dir, file.filename)
        file.save(filepath)
        try:
            from .services.excelloader import load_excel, _debug_matrix
            results   = load_excel(filepath)
            score_msg = (f"{results.get('scores',0)} scores"
                         if results.get('score_row_found')
                         else '⚠ score row not detected')
            flash(
                f"Loaded: {results['models']} models, {results['parameters']} parameters, "
                f"{score_msg}, {results['tiers']} tiers. "
                f"Tiering computed for {results.get('computed', 0)} models.",
                'success'
            )
            if not results.get('score_row_found') or results['tiers'] == 0:
                try:
                    dbg = _debug_matrix(filepath)
                    flash(f'DEBUG: {dbg}', 'warning')
                except Exception:
                    pass
        except Exception as e:
            import traceback
            flash(f'Error loading Excel: {str(e)}', 'danger')
            flash(f'Trace: {traceback.format_exc()[:600]}', 'warning')
        return redirect(url_for('main.upload'))
    return render_template('upload.html')


# ── Settings ──────────────────────────────────────────────────────────────────

@main_bp.route('/settings', methods=['GET', 'POST'])
@login_required
@roles_required('Admin')
def settings():
    db = get_db()

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'weights':
            for key in ('materiality_weight', 'criticality_weight', 'complexity_weight'):
                val = request.form.get(key, '')
                db.execute('INSERT OR REPLACE INTO config_kv (key, value) VALUES (?,?)', (key, val))
            db.commit()
            compute_tiering_for_all()
            flash('Weights updated and scores recomputed.', 'success')

        elif action == 'tier':
            tid = request.form.get('tier_id')
            db.execute('UPDATE tiers SET name=?, lower_bound=?, upper_bound=? WHERE id=?',
                       (request.form.get('tier_name'),
                        float(request.form.get('lower_bound', 0)),
                        float(request.form.get('upper_bound', 100)), tid))
            db.commit()
            flash('Tier updated.', 'success')

        elif action == 'add_tier':
            sort = db.execute('SELECT COUNT(*) FROM tiers').fetchone()[0]
            db.execute('INSERT INTO tiers (name, lower_bound, upper_bound, sort_order) VALUES (?,?,?,?)',
                       (request.form.get('tier_name', 'New Tier'),
                        float(request.form.get('lower_bound', 0)),
                        float(request.form.get('upper_bound', 100)), sort))
            db.commit()
            flash('Tier added.', 'success')

        elif action == 'delete_tier':
            db.execute('DELETE FROM tiers WHERE id=?', (request.form.get('tier_id'),))
            db.commit()
            flash('Tier deleted.', 'success')

        elif action == 'reset_ids':
            db.execute('DELETE FROM model_scores')
            db.execute('DELETE FROM overrides')
            db.execute('DELETE FROM models')
            db.execute('DELETE FROM parameters')
            db.execute("DELETE FROM sqlite_sequence WHERE name IN "
                       "('models','parameters','model_scores','overrides')")
            db.commit()
            flash('Database reset. Re-upload your Excel to start fresh.', 'success')

        return redirect(url_for('main.settings'))

    tiers = db.execute('SELECT * FROM tiers ORDER BY sort_order').fetchall()
    users = db.execute('SELECT * FROM users ORDER BY email').fetchall()

    def gcfg(key, default):
        row = db.execute('SELECT value FROM config_kv WHERE key=?', (key,)).fetchone()
        return row['value'] if row else default

    return render_template('settings.html',
                           tiers=tiers, users=users,
                           mat_w=gcfg('materiality_weight', '0.4'),
                           crit_w=gcfg('criticality_weight', '0.4'),
                           comp_w=gcfg('complexity_weight',  '0.2'))
