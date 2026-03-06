from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from . import db


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(50), nullable=False, default='User')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.email}>'


class Model(db.Model):
    __tablename__ = 'models'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    risk_type = db.Column(db.String(100))
    computed_score = db.Column(db.Float, default=0.0)
    computed_tier = db.Column(db.String(50))
    current_tier = db.Column(db.String(50))
    last_computed_at = db.Column(db.DateTime)

    scores = db.relationship('ModelScore', backref='model', lazy=True, cascade='all, delete-orphan')
    overrides = db.relationship('Override', backref='model', lazy=True, cascade='all, delete-orphan')

    def __repr__(self):
        return f'<Model {self.name}>'


class Parameter(db.Model):
    __tablename__ = 'parameters'
    id = db.Column(db.Integer, primary_key=True)
    group = db.Column(db.String(100), nullable=False)
    sub_parameter = db.Column(db.String(255))
    criteria = db.Column(db.String(255))
    description = db.Column(db.Text)
    weight = db.Column(db.Float, default=1.0)

    scores = db.relationship('ModelScore', backref='parameter', lazy=True, cascade='all, delete-orphan')

    def __repr__(self):
        return f'<Parameter {self.group} - {self.sub_parameter}>'


class ModelScore(db.Model):
    __tablename__ = 'model_scores'
    id = db.Column(db.Integer, primary_key=True)
    model_id = db.Column(db.Integer, db.ForeignKey('models.id'), nullable=False)
    parameter_id = db.Column(db.Integer, db.ForeignKey('parameters.id'), nullable=False)
    level = db.Column(db.Integer, default=1)  # 1=Low, 2=Medium, 3=High
    weighted_score = db.Column(db.Float, default=0.0)


class Tier(db.Model):
    __tablename__ = 'tiers'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    lower_bound = db.Column(db.Float, nullable=False)
    upper_bound = db.Column(db.Float, nullable=False)
    sort_order = db.Column(db.Integer, default=0)

    def __repr__(self):
        return f'<Tier {self.name}>'


class Override(db.Model):
    __tablename__ = 'overrides'
    id = db.Column(db.Integer, primary_key=True)
    model_id = db.Column(db.Integer, db.ForeignKey('models.id'), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    old_tier = db.Column(db.String(50))
    new_tier = db.Column(db.String(50))
    reason = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref='overrides')


class Report(db.Model):
    __tablename__ = 'reports'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    name = db.Column(db.String(255), nullable=False)
    filters_json = db.Column(db.Text, default='{}')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref='reports')


class ConfigKV(db.Model):
    __tablename__ = 'config_kv'
    key = db.Column(db.String(255), primary_key=True)
    value = db.Column(db.Text)

    @classmethod
    def get(cls, key, default=None):
        row = cls.query.filter_by(key=key).first()
        return row.value if row else default

    @classmethod
    def set(cls, key, value):
        row = cls.query.filter_by(key=key).first()
        if row:
            row.value = str(value)
        else:
            row = cls(key=key, value=str(value))
            db.session.add(row)
        db.session.commit()
