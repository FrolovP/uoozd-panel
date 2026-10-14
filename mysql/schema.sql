-- ============================================================
-- Панель УООЗД — схема MySQL 8.0+
-- Кодировка utf8mb4, движок InnoDB, внешние ключи включены.
-- Применение:  mysql -u root -p < schema.sql
-- ============================================================

CREATE DATABASE IF NOT EXISTS uoozd
  DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE uoozd;

-- ------------------------------------------------------------
-- ПОЛЬЗОВАТЕЛИ, СЕССИИ, ВХОД
-- ------------------------------------------------------------

CREATE TABLE users (
  id            INT UNSIGNED   NOT NULL AUTO_INCREMENT PRIMARY KEY,
  login         VARCHAR(32)    NOT NULL,
  fio           VARCHAR(64)    NOT NULL,
  dept          ENUM('ОАО','ОПАиЗ','ОВРК','УООЗД') NOT NULL DEFAULT 'ОАО',
  is_admin      TINYINT(1)     NOT NULL DEFAULT 0,
  can_delete    TINYINT(1)     NOT NULL DEFAULT 0,
  must_change   TINYINT(1)     NOT NULL DEFAULT 1,
  is_active     TINYINT(1)     NOT NULL DEFAULT 1,
  salt          CHAR(32)       NOT NULL,               -- hex, 16 байт
  pass_hash     CHAR(64)       NOT NULL,               -- PBKDF2-HMAC-SHA256, hex
  hash_iters    INT UNSIGNED   NOT NULL DEFAULT 210000,
  failed_logins INT UNSIGNED   NOT NULL DEFAULT 0,
  locked_until  DATETIME       NULL,
  pwd_changed_at DATETIME      NULL,
  created_at    DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at    DATETIME       NOT NULL DEFAULT CURRENT_TIMESTAMP
                               ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uq_users_login (login)
) ENGINE=InnoDB;

CREATE TABLE sessions (
  token_hash  CHAR(64)      NOT NULL PRIMARY KEY,      -- sha256 от токена, сам токен в БД не храним
  user_id     INT UNSIGNED  NOT NULL,
  ip          VARBINARY(16) NULL,
  user_agent  VARCHAR(255)  NULL,
  created_at  DATETIME      NOT NULL DEFAULT CURRENT_TIMESTAMP,
  expires_at  DATETIME      NOT NULL,
  revoked     TINYINT(1)    NOT NULL DEFAULT 0,
  KEY ix_sessions_user (user_id),
  KEY ix_sessions_exp  (expires_at),
  CONSTRAINT fk_sessions_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE login_attempts (
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  login       VARCHAR(32)  NOT NULL,
  ip          VARBINARY(16) NULL,
  ok          TINYINT(1)   NOT NULL,
  at          DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  KEY ix_attempts_login_at (login, at),
  KEY ix_attempts_ip_at (ip, at)
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- СПРАВОЧНИКИ
-- ------------------------------------------------------------

CREATE TABLE departments (
  id    SMALLINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  name  VARCHAR(128) NOT NULL UNIQUE
) ENGINE=InnoDB;

CREATE TABLE employees (
  id         INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  fio        VARCHAR(64)  NOT NULL,                    -- «Фамилия И.О.»
  email      VARCHAR(128) NULL,
  dept       ENUM('ОАО','ОПАиЗ','ОВРК','УООЗД') NOT NULL,
  role_title VARCHAR(64)  NULL,                        -- «начальник отдела» и т.п.
  is_active  TINYINT(1)   NOT NULL DEFAULT 1,
  UNIQUE KEY uq_emp_fio_dept (fio, dept)
) ENGINE=InnoDB;

CREATE TABLE counterparties (
  id        INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  name      VARCHAR(255) NOT NULL,
  inn       VARCHAR(12)  NOT NULL,
  kpp       VARCHAR(9)   NULL,
  ogrn      VARCHAR(15)  NULL,
  address   VARCHAR(255) NULL,
  email     VARCHAR(128) NULL,
  smp_cat   ENUM('микро','малое','среднее','СОНКО','') NOT NULL DEFAULT '',
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_counterparty_inn (inn),
  KEY ix_counterparty_name (name)
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- РЕЕСТР ОАО: закупки, протоколы, договоры
-- ------------------------------------------------------------

CREATE TABLE procurements (
  id            BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  year          SMALLINT        NOT NULL,              -- год реестра
  sheet         ENUM('k','nk')  NOT NULL,              -- конкурентные / неконкурентные
  oao_number    VARCHAR(32)     NOT NULL,              -- № ОАО (1535зэа)
  pub_number    VARCHAR(32)     NULL,                  -- № на площадке
  subject       TEXT            NOT NULL,
  method        VARCHAR(64)     NULL,                  -- способ закупки
  nmck          DECIMAL(18,2)   NULL,
  responsible   VARCHAR(64)     NULL,                  -- «Фамилия И.О.»
  platform      VARCHAR(128)    NULL,
  dt_publish    DATE NULL, dt_open DATE NULL, dt_review DATE NULL,
  dt_results    DATE NULL, dt_contract_plan DATE NULL,
  winner_id     INT UNSIGNED    NULL,
  winner_price  DECIMAL(18,2)   NULL,
  economy       DECIMAL(18,2)   NULL,
  status        VARCHAR(32)     NOT NULL DEFAULT 'в работе',
  smp_cat       VARCHAR(16)     NULL,                  -- «Кат. см и сп»
  is_science    TINYINT(1)      NOT NULL DEFAULT 0,    -- «★ Наука»
  deleted       TINYINT(1)      NOT NULL DEFAULT 0,
  delete_reason VARCHAR(255)    NULL,
  payload       JSON            NULL,                  -- поля, которых нет колонками (не ломать импорт)
  created_at    DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at    DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uq_proc_year_oao (year, sheet, oao_number),
  KEY ix_proc_pub   (pub_number),
  KEY ix_proc_resp  (responsible),
  KEY ix_proc_dates (dt_results),
  KEY ix_proc_winner (winner_id),
  CONSTRAINT fk_proc_winner FOREIGN KEY (winner_id) REFERENCES counterparties(id)
) ENGINE=InnoDB;

CREATE TABLE protocols (
  id           BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  procurement_id BIGINT UNSIGNED NOT NULL,
  proto_number SMALLINT NOT NULL,                      -- протокол №1, №2…
  proto_kind   VARCHAR(32)  NULL,                      -- вскрытие / рассмотрение / итоги
  proto_date   DATE NULL,
  source_file  VARCHAR(255) NULL,                      -- относительный путь docx/doc/xml
  source_sha256 CHAR(64)    NULL,
  payload      JSON         NULL,
  created_at   DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_proto (procurement_id, proto_number, proto_kind),
  CONSTRAINT fk_proto_proc FOREIGN KEY (procurement_id) REFERENCES procurements(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE protocol_participants (
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  protocol_id BIGINT UNSIGNED NOT NULL,
  counterparty_id INT UNSIGNED NULL,
  name_raw    VARCHAR(255) NOT NULL,                   -- как в файле
  inn         VARCHAR(12)  NULL,
  price       DECIMAL(18,2) NULL,
  score       DECIMAL(10,3)  NULL,                     -- баллы (конкурс)
  is_winner   TINYINT(1) NOT NULL DEFAULT 0,
  rejected    TINYINT(1) NOT NULL DEFAULT 0,
  reject_reason VARCHAR(255) NULL,
  KEY ix_part_inn (inn),
  CONSTRAINT fk_part_proto FOREIGN KEY (protocol_id) REFERENCES protocols(id) ON DELETE CASCADE,
  CONSTRAINT fk_part_cp FOREIGN KEY (counterparty_id) REFERENCES counterparties(id)
) ENGINE=InnoDB;

CREATE TABLE contracts (
  id             BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  procurement_id BIGINT UNSIGNED NULL,                 -- NULL — прямой договор без процедуры
  number         VARCHAR(64) NOT NULL,
  dt_conclusion  DATE NULL,
  dt_exec_from   DATE NULL, dt_exec_to DATE NULL,      -- «с … по …» из текста
  dt_valid_until DATE NULL,
  price          DECIMAL(18,2) NULL,
  counterparty_id INT UNSIGNED NULL,
  source_file    VARCHAR(255) NULL,
  source_sha256  CHAR(64) NULL,
  payload        JSON NULL,
  created_at     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uq_contract (number, dt_conclusion),
  CONSTRAINT fk_contract_proc FOREIGN KEY (procurement_id) REFERENCES procurements(id),
  CONSTRAINT fk_contract_cp   FOREIGN KEY (counterparty_id) REFERENCES counterparties(id)
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- ОВРК: исполнение договоров
-- ------------------------------------------------------------

CREATE TABLE ovrk_documents (                          -- УПД, акты
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  contract_id BIGINT UNSIGNED NOT NULL,
  doc_number  VARCHAR(64) NOT NULL,
  doc_date    DATE NOT NULL,
  amount      DECIMAL(18,2) NOT NULL,
  responsible VARCHAR(64) NULL,                        -- ответственный за размещение в ЕИС
  eis_date    DATE NULL,                               -- когда внесён в ЕИС
  note        VARCHAR(500) NULL,
  file_path   VARCHAR(255) NULL,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uq_ovrk_doc (contract_id, doc_number, doc_date),
  CONSTRAINT fk_ovrkdoc_contract FOREIGN KEY (contract_id) REFERENCES contracts(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE ovrk_payments (                           -- платёжки (ПП)
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  document_id BIGINT UNSIGNED NOT NULL,
  pp_number   VARCHAR(32) NOT NULL,
  pp_date     DATE NOT NULL,
  amount      DECIMAL(18,2) NOT NULL,
  payee_inn   VARCHAR(12) NULL,                        -- ИНН получателя из файла платёжки
  purpose     TEXT NULL,                               -- назначение платежа
  eis_status  ENUM('нет','внесена') NOT NULL DEFAULT 'нет',
  file_path   VARCHAR(255) NULL,
  file_sha256 CHAR(64) NULL,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_pp (pp_number, pp_date),
  KEY ix_pp_doc (document_id),
  CONSTRAINT fk_pp_doc FOREIGN KEY (document_id) REFERENCES ovrk_documents(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE contract_addenda (                        -- доп. соглашения
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  contract_id BIGINT UNSIGNED NOT NULL,
  number      VARCHAR(64) NOT NULL,
  ds_date     DATE NULL,
  reason      VARCHAR(255) NULL,
  new_price   DECIMAL(18,2) NULL,
  new_exec_to DATE NULL,
  new_valid_until DATE NULL,
  CONSTRAINT fk_ds_contract FOREIGN KEY (contract_id) REFERENCES contracts(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE claims (                                  -- претензии
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  contract_id BIGINT UNSIGNED NOT NULL,
  number      VARCHAR(64) NULL,
  claim_date  DATE NULL,
  amount      DECIMAL(18,2) NULL,
  basis       VARCHAR(255) NULL,
  status      VARCHAR(32) NULL,
  reply_until DATE NULL,
  CONSTRAINT fk_claim_contract FOREIGN KEY (contract_id) REFERENCES contracts(id) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- ОПАиЗ
-- ------------------------------------------------------------

CREATE TABLE oz_requests (                             -- электронные заявки
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  num_1c      VARCHAR(32) NULL,                        -- № заявки 1С — связь с реестром ОАО
  nmck        DECIMAL(18,2) NULL,
  subject     TEXT NULL,
  initiator   VARCHAR(64) NULL,
  subdivision VARCHAR(128) NULL,
  do_stage    VARCHAR(64) NULL,                        -- согласование в 1С ДО: этап
  do_holder   VARCHAR(64) NULL,                        -- у кого
  do_since    DATE NULL,
  status      VARCHAR(32) NOT NULL DEFAULT 'в работе',
  procurement_id BIGINT UNSIGNED NULL,                 -- подтянутая закупка ОАО
  payload     JSON NULL,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uq_oz_num1c (num_1c),
  KEY ix_oz_status (status),
  CONSTRAINT fk_oz_proc FOREIGN KEY (procurement_id) REFERENCES procurements(id)
) ENGINE=InnoDB;

CREATE TABLE oz_memos (                                -- служебные записки СЗ 1С
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  number      VARCHAR(32) NOT NULL,
  priority    VARCHAR(32) NULL,                        -- приоритет (паспорт)
  prorector   VARCHAR(64) NULL,
  status      ENUM('в работе','исполнена','отклонена') NOT NULL DEFAULT 'в работе',
  complexity  TINYINT NULL, urgency TINYINT NULL,      -- ставит начальник отдела
  deadline    DATE NULL,
  assignee    VARCHAR(64) NULL,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_memo (number)
) ENGINE=InnoDB;

CREATE TABLE oz_memo_items (                           -- позиции СЗ (потребность)
  id        BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  memo_id   BIGINT UNSIGNED NOT NULL,
  category  VARCHAR(64) NULL,                          -- лист Excel / словарь категорий
  name      VARCHAR(500) NOT NULL,
  okpd2     VARCHAR(32) NULL,
  qty       DECIMAL(18,3) NULL,
  unit      VARCHAR(32) NULL,
  is_rf_equipment TINYINT(1) NOT NULL DEFAULT 0,
  assignee  VARCHAR(64) NULL,                          -- распределение по листам
  KEY ix_items_memo (memo_id),
  CONSTRAINT fk_item_memo FOREIGN KEY (memo_id) REFERENCES oz_memos(id) ON DELETE CASCADE
) ENGINE=InnoDB;

CREATE TABLE oz_kp_requests (                          -- запросы коммерческих предложений
  id        BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  incoming_no VARCHAR(32) NULL,                        -- «15/10-2026», ставится сам
  subject   TEXT NULL,
  cp_price  DECIMAL(18,2) NULL,
  received_at DATE NULL,
  payload   JSON NULL
) ENGINE=InnoDB;

CREATE TABLE oz_ep_contracts (                         -- договоры с единственным поставщиком
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  contract_id BIGINT UNSIGNED NULL,                    -- если договор заведён в ОВРК
  number      VARCHAR(64) NULL,
  ep_date     DATE NULL,
  subject     TEXT NULL,
  price       DECIMAL(18,2) NULL,
  advance_pct TINYINT NULL,
  smp_check   VARCHAR(255) NULL,                       -- итог проверки СМП
  rf_equipment TINYINT(1) NOT NULL DEFAULT 0,
  payload     JSON NULL,
  CONSTRAINT fk_ep_contract FOREIGN KEY (contract_id) REFERENCES contracts(id)
) ENGINE=InnoDB;

CREATE TABLE oz_invoices (                             -- счета
  id        BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  number    VARCHAR(64) NULL,
  inv_date  DATE NULL,
  amount    DECIMAL(18,2) NULL,
  pay_terms VARCHAR(64) NULL,                          -- аванс 5 раб.дн. / постоплата 7 раб.дн.
  delivery_terms VARCHAR(64) NULL,
  handed_to_oao DATE NULL,                             -- передан в ОАО (контроль 50 раб.дн.)
  payload   JSON NULL
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- ИМПОРТ ФАЙЛОВ (Excel и XML)
-- ------------------------------------------------------------

CREATE TABLE import_batches (
  id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  source      ENUM('excel','xml','folder','panel') NOT NULL,
  file_name   VARCHAR(255) NOT NULL,
  file_sha256 CHAR(64) NOT NULL,
  format_key  VARCHAR(64) NULL,                        -- какой формат распознан (см. базу знаний)
  status      ENUM('новый','разобран','на проверке','отклонён','ошибка') NOT NULL DEFAULT 'новый',
  message     VARCHAR(500) NULL,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_batch_file (file_name, file_sha256)    -- повтор того же файла не вносится
) ENGINE=InnoDB;

CREATE TABLE import_items (
  id        BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id  BIGINT UNSIGNED NOT NULL,
  item_kind VARCHAR(32) NOT NULL,                      -- извещение / протокол / договор / платёжка …
  fields    JSON NOT NULL,                             -- распознанные поля до записи в реестр
  applied   TINYINT(1) NOT NULL DEFAULT 0,
  CONSTRAINT fk_item_batch FOREIGN KEY (batch_id) REFERENCES import_batches(id) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- ЖУРНАЛЫ
-- ------------------------------------------------------------

CREATE TABLE audit_log (                               -- заменяет «Реестр — журнал.json», append-only
  id      BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  ts      DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
  kind    VARCHAR(16) NOT NULL,                        -- info / edit / error / auth
  user_fio VARCHAR(64) NULL,
  ip      VARBINARY(16) NULL,
  message TEXT NOT NULL,
  KEY ix_audit_ts (ts), KEY ix_audit_kind (kind)
) ENGINE=InnoDB;

CREATE TABLE news (
  id       INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  title    VARCHAR(200) NOT NULL,
  body     TEXT NOT NULL,
  dept     VARCHAR(16) NULL,
  valid_from DATE NULL,
  important TINYINT(1) NOT NULL DEFAULT 0,
  author   VARCHAR(64) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE notifications (
  id        BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  channel   ENUM('макс','почта') NOT NULL,
  event     VARCHAR(64) NOT NULL,
  address   VARCHAR(128) NULL,
  body      TEXT NULL,
  status    VARCHAR(32) NOT NULL,                      -- отправлено / отправлено (демо) / ошибка
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  KEY ix_notify_channel (channel)
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- БАЗА ЗНАНИЙ
-- ------------------------------------------------------------

CREATE TABLE kb_articles (
  id        INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  slug      VARCHAR(64) NOT NULL UNIQUE,               -- 'zayavka-excel', 'izv-xml' …
  title     VARCHAR(200) NOT NULL,
  category  VARCHAR(64) NOT NULL,                      -- Заявки / Извещения / Протоколы / Договоры / Платёжки / XML
  body_md   MEDIUMTEXT NOT NULL,                       -- Markdown, рендерится в разделе «База знаний»
  updated_by VARCHAR(64) NULL,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ------------------------------------------------------------
-- СЛУЖЕБНОЕ
-- ------------------------------------------------------------

CREATE TABLE settings (                                -- СГОЗ, email бухгалтерии, ПРАЗДНИКИ и т.п.
  skey   VARCHAR(64) NOT NULL PRIMARY KEY,
  svalue JSON NULL
) ENGINE=InnoDB;

-- Отдельный пользователь для приложения (пароль задать при установке):
-- CREATE USER 'uoozd_app'@'%' IDENTIFIED BY '<пароль из сейфа>';
-- GRANT SELECT, INSERT, UPDATE, DELETE ON uoozd.* TO 'uoozd_app'@'%';
-- DELETE на audit_log и login_attempts НЕ давать (журнал неизменяем):
-- REVOKE DELETE ON uoozd.audit_log FROM 'uoozd_app'@'%';  -- либо права только перечнем таблиц
