/* ===============================
           USER CREATION
   =============================== */

DROP USER PRJ2531G CASCADE;

CREATE USER PRJ2531G IDENTIFIED BY PRJ2531G;
GRANT CONNECT, RESOURCE TO PRJ2531G;
GRANT DBA TO PRJ2531G;

CONN PRJ2531G/PRJ2531G;


/* ===============================
        USER DETAILS (LOGIN)
   =============================== */

CREATE TABLE USER_DETAILS (
    usernm      VARCHAR2(30) PRIMARY KEY,
    pass        VARCHAR2(30) NOT NULL,
    fullname    VARCHAR2(50),
    sec_ques    VARCHAR2(100),
    sec_ans     VARCHAR2(50)
);


/* ===============================
         1. REF_BLOOD_RANGE
   =============================== */

CREATE TABLE REF_BLOOD_RANGE (
  ref_id   VARCHAR2(10) CONSTRAINT REF_ID_PK PRIMARY KEY,
  cm_nm    VARCHAR2(50) NOT NULL,
  me_un    VARCHAR2(20) NOT NULL,
  mn_v_m   VARCHAR2(20),
  mx_v_m   VARCHAR2(20),
  mn_v_f   VARCHAR2(20),
  mx_v_f   VARCHAR2(20),
  age      VARCHAR2(20),
  descr    VARCHAR2(100),
  status   VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
        2. DONAR DETAIL
   =============================== */

CREATE TABLE DONAR_DETAIL (
  d_id       VARCHAR2(10) CONSTRAINT D_ID_PK PRIMARY KEY,
  nm         VARCHAR2(25) NOT NULL,
  f_nm       VARCHAR2(25) NOT NULL,
  dob        DATE NOT NULL,
  gender     VARCHAR2(6) CHECK (gender IN ('Male','Female')),
  bl_grp     VARCHAR2(3) CHECK (bl_grp IN ('A+','B+','AB+','O+','A-','B-','AB-','O-')),
  occ        VARCHAR2(20),
  org        VARCHAR2(25),
  addr       VARCHAR2(25),
  pin        NUMBER(6),
  po         VARCHAR2(10),
  dist       VARCHAR2(10),
  tele       NUMBER(15),
  phn        NUMBER(10) UNIQUE,
  call       CHAR(3) CHECK (call IN ('Yes','No')),
  fax        VARCHAR2(20),
  email      VARCHAR2(100) UNIQUE,
  d_prev     VARCHAR2(3) CHECK (d_prev IN ('YES','NO')),
  tot_don    NUMBER(3) DEFAULT 0,
  last_don   DATE,
  status     VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
        3. DONATION VISIT
   =============================== */

CREATE TABLE DONATION_VISIT (
  v_id      VARCHAR2(10) CONSTRAINT v_id_PK PRIMARY KEY,
  d_id      VARCHAR2(10) CONSTRAINT d_id_FK REFERENCES DONAR_DETAIL(d_id),
  don_dt    DATE NOT NULL,
  ref_nm    VARCHAR2(25),
  opd       NUMBER(5),
  bed       VARCHAR2(5),
  wgt       NUMBER(5,2) NOT NULL,
  bp        VARCHAR2(7),
  pulse     NUMBER(3),
  tem       NUMBER(4,1),
  hb_lev    NUMBER(4,1),
  fin_stat  VARCHAR2(10) CHECK (fin_stat IN ('FIT','UNFIT','DEFERRED')),
  defer     VARCHAR2(150),
  status    VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
          4. CONSENT
   =============================== */

CREATE TABLE CONSENT (
  co_id    VARCHAR2(10) CONSTRAINT CO_ID_PK PRIMARY KEY,
  v_id     VARCHAR2(10) CONSTRAINT v_id_FK REFERENCES DONATION_VISIT(v_id),
  dis      VARCHAR2(25),
  well     CHAR(1) DEFAULT 'N' CHECK (well IN ('Y','N')),
  eat      CHAR(1) DEFAULT 'N' CHECK (eat IN ('Y','N')),
  sleep    CHAR(1) DEFAULT 'N' CHECK (sleep IN ('Y','N')),
  infect   CHAR(1) DEFAULT 'N' CHECK (infect IN ('Y','N')),
  his      VARCHAR2(25),
  last_6   CHAR(1) DEFAULT 'N' CHECK (last_6 IN ('Y','N')),
  preg     CHAR(1) DEFAULT 'N' CHECK (preg IN ('Y','N')),
  abor     CHAR(1) DEFAULT 'N' CHECK (abor IN ('Y','N')),
  child    CHAR(1) DEFAULT 'N' CHECK (child IN ('Y','N')),
  abnor    CHAR(1) DEFAULT 'N' CHECK (abnor IN ('Y','N'))
);


/* ===============================
          5. DONAR DISEASE
   =============================== */

CREATE TABLE D_DISEASE (
  dis_id  VARCHAR2(10) CONSTRAINT DIS_ID_PK PRIMARY KEY,
  v_id    VARCHAR2(10) CONSTRAINT dise_vid_FK REFERENCES DONATION_VISIT(v_id),
  dise    VARCHAR2(50),
  hr_72   CHAR(1) DEFAULT 'N' CHECK (hr_72 IN ('Y','N')),
  surg    CHAR(1) DEFAULT 'N' CHECK (surg IN ('Y','N'))
);


/* ===============================
          6. SAMPLE
   =============================== */

CREATE TABLE SAMPLE (
  s_id    VARCHAR2(10) CONSTRAINT S_ID_PK PRIMARY KEY,
  v_id    VARCHAR2(10) CONSTRAINT sam_vid_FK REFERENCES DONATION_VISIT(v_id),
  c_nm    VARCHAR2(30) NOT NULL,
  u_id    VARCHAR2(20) UNIQUE NOT NULL,
  c_date  DATE NOT NULL,
  status  VARCHAR2(20),
  status2 VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
          7. DONAR CARD
   =============================== */

CREATE TABLE DONAR_CARD (
  c_no    NUMBER(10) CONSTRAINT C_NO_PK PRIMARY KEY,
  d_id    VARCHAR2(10) CONSTRAINT card_did_FK REFERENCES DONAR_DETAIL(d_id),
  iss_dt  DATE NOT NULL,
  valid   DATE NOT NULL,
  status  VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
          8. LAB TESTING
   =============================== */

CREATE TABLE LAB_TESTING (
  t_id     VARCHAR2(10) CONSTRAINT T_ID_PK PRIMARY KEY,
  s_id     VARCHAR2(10) CONSTRAINT lab_sid_FK REFERENCES SAMPLE(s_id),
  hiv_s    VARCHAR2(3) DEFAULT 'NEG' CHECK (hiv_s IN ('POS','NEG')),
  hcv_s    VARCHAR2(3) DEFAULT 'NEG' CHECK (hcv_s IN ('POS','NEG')),
  mala     VARCHAR2(3) DEFAULT 'NEG' CHECK (mala IN ('POS','NEG')),
  fin_res  VARCHAR2(4) CHECK (fin_res IN ('PASS','FAIL')),
  status   VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
       9. COMPONENT PROCESSING
   =============================== */

CREATE TABLE COMPONENT_PROCESSING (
  p_id    VARCHAR2(10) CONSTRAINT P_ID_PK PRIMARY KEY,
  t_id    VARCHAR2(10) CONSTRAINT comp_tid_FK REFERENCES LAB_TESTING(t_id),
  com_nm  VARCHAR2(50) NOT NULL,
  un_id   VARCHAR2(20) UNIQUE,
  exp     DATE NOT NULL,
  vol     NUMBER(6,2),
  status  VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
          10. BLOOD STOCK
   =============================== */

CREATE TABLE BLOOD_STOCK (
  st_id   VARCHAR2(10) CONSTRAINT ST_ID_PK PRIMARY KEY,
  p_id    VARCHAR2(10) CONSTRAINT sto_st_id_FK REFERENCES COMPONENT_PROCESSING(p_id),
  rack    VARCHAR2(10),
  avail   VARCHAR2(10) CHECK (avail IN ('AVAILABLE','ISSUED','EXPIRED')),
  status  VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
          11. PATIENT
   =============================== */

CREATE TABLE PATIENT (
  pa_id   VARCHAR2(10) CONSTRAINT PA_ID_PK PRIMARY KEY,
  p_nm    VARCHAR2(50) NOT NULL,
  p_grp   VARCHAR2(3) CHECK (p_grp IN ('A+','B+','AB+','O+','A-','B-','AB-','O-')),
  addr    VARCHAR2(200),
  p_pn    NUMBER(10),
  status  VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
         12. PATIENT REQUEST
   =============================== */

CREATE TABLE PATIENT_REQUEST (
  req_id  VARCHAR2(10) CONSTRAINT REQ_ID_PK PRIMARY KEY,
  pa_id   VARCHAR2(10) CONSTRAINT pat_pa_id_FK REFERENCES PATIENT(pa_id),
  hos_nm  VARCHAR2(20),
  wrd     VARCHAR2(10),
  bed     VARCHAR2(10),
  units   NUMBER(3) NOT NULL,
  comp    VARCHAR2(50) NOT NULL,
  hb_p    NUMBER(6,2),
  pre     VARCHAR2(15),
  plat    NUMBER(6),
  dt      DATE NOT NULL,
  d_nm    VARCHAR2(25),
  status  VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
         13. BLOOD ISSUE
   =============================== */

CREATE TABLE BLOOD_ISSUE (
  issue_id  VARCHAR2(10) CONSTRAINT ISSUE_ID_PK PRIMARY KEY,
  req_id    VARCHAR2(10) CONSTRAINT iss_req_id_FK REFERENCES PATIENT_REQUEST(req_id),
  st_id     VARCHAR2(10) REFERENCES BLOOD_STOCK(st_id),
  x_match   VARCHAR2(20) CHECK (x_match IN ('MATCH','NOT MATCH')),
  issue_dt  DATE NOT NULL,
  issue_by  VARCHAR2(25),
  recv_by   VARCHAR2(25),
  fee       NUMBER(8,2),
  bill      NUMBER(6),
  mod       VARCHAR2(20),
  pay       VARCHAR2(10) CHECK (pay IN ('PAID','PENDING')),
  status    VARCHAR2(10) DEFAULT 'Active'
);


/* ===============================
          FINAL MESSAGE
   =============================== */

SET HEADING OFF
SET FEEDBACK OFF
SET PAGESIZE 0

SELECT 'All Tables Successfully Created' FROM DUAL;
SELECT 'PROJECT ID - PRJ2531G' FROM DUAL;
SELECT 'BLOOD BANK MANAGEMENT SYSTEM' FROM DUAL;

SET HEADING ON
SET FEEDBACK ON