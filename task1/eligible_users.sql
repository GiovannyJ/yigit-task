CREATE TABLE dbo.users(
    emplid INT NOT NULL AUTO_INCREMENT,
    fullname CHAR(25) NOT NULL,
    fname CHAR(25) NOT NULL,
    lname CHAR(25) NOT NULL,
    status NUMBER(1) NOT NULL DEFAULT 0,
    position CHAR(23) NOT NULL,
    ssn INT NOT NULL,
    eligible NUMBER(1) NOT NULL DEFAULT 0,
    title CHAR(20) NOT NULL,
    dept CHAR(30) NOT NULL,
    depnum INT NOT NULL,
    deleted NUMBER(1) NOT NULL DEFAULT 0
);
[rs1, rs2, rs3]
INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (1, 'John Doe', 'John', 'Doe', 1, 'Nurse', 123456789, 1, 'RN', 'Nursing', 10, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (2, 'Jane Smith', 'Jane', 'Smith', 1, 'Doctor', 234567890, 1, 'MD', 'Surgery', 20, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (3, 'Jim Brown', 'Jim', 'Brown', 1, 'Receptionist', 345678901, 1, 'Receptionist', 'Administration', 30, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (4, 'Emily Davis', 'Emily', 'Davis', 1, 'Pharmacist', 456789012, 1, 'PharmD', 'Pharmacy', 40, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (5, 'Michael Wilson', 'Michael', 'Wilson', 1, 'Lab Technician', 567890123, 1, 'Lab Tech', 'Laboratory', 50, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (6, 'Sara Lee', 'Sara', 'Lee', 1, 'Radiologist', 678901234, 1, 'MD', 'Radiology', 60, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (7, 'William Green', 'William', 'Green', 1, 'Surgeon', 789012345, 1, 'MD', 'Surgery', 20, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (8, 'Emily Clark', 'Emily', 'Clark', 1, 'Nurse Practitioner', 890123456, 1, 'NP', 'Nursing', 10, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (9, 'Daniel Anderson', 'Daniel', 'Anderson', 1, 'Anesthesiologist', 901234567, 1, 'MD', 10, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (10, 'Ashley Johnson', 'Ashley', 'Johnson', 1, 'Physical Therapist', 012345678, 1, 'PT', 'Rehabilitation', 70, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (11, 'Michael Martinez', 'Michael', 'Martinez', 1, 'Cardiologist', 123456789, 1, 'MD', 'Cardiology', 80, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (12, 'Daniel Lee', 'Daniel', 'Lee', 1, 'Oncologist', 234567890, 1, 'MD', 'Oncology', 90, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (13, 'Emily Gonzalez', 'Emily', 'Gonzalez', 1, 'Optometrist', 345678901, 1, 'OD', 'Optometry', 100, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (14, 'John King', 'John', 'King', 1, 'Neurologist', 456789012, 1, 'MD', 'Neurology', 110, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (15, 'William Brown', 'William', 'Brown', 1, 'Dentist', 567890123, 1, 'DDS', 'Dentistry', 120, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (16, 'Jane Green', 'Jane', 'Green', 1, 'Pediatrician', 678901234, 1, 'MD', 'Pediatrics', 130, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (17, 'Jim Anderson', 'Jim', 'Anderson', 1, 'Dermatologist', 789012345, 1, 'MD', 'Dermatology', 140, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (18, 'Michael Davis', 'Michael', 'Davis', 1, 'Orthopedic Surgeon', 890123456, 1, 'MD', 'Orthopedics', 150, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (19, 'Emily Kim', 'Emily', 'Kim', 1, 'Gastroenterologist', 901234567, 1, 'MD', 'Gastroenterology', 160, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (20, 'William Wilson', 'William', 'Wilson', 1, 'Gynecologist', 012345678, 1, 'MD', 'Gynecology', 170, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (21, 'Daniel Lee', 'Daniel', 'Lee', 1, 'Anesthesiologist', 123456789, 1, 'MD', 'Anesthesiology', 180, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (22, 'John Smith', 'John', 'Smith', 1, 'Ophthalmologist', 234567890, 1, 'MD', 'Ophthalmology', 190, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (23, 'Jane Kim', 'Jane', 'Kim', 1, 'Podiatrist', 345678901, 1, 'DPM', 'Podiatry', 200, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (24, 'James Wilson', 'James', 'Wilson', 1, 'Psychiatrist', 456789012, 1, 'MD', 'Psychiatry', 210, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (25, 'Michael Davis', 'Michael', 'Davis', 1, 'Radiologist', 567890123, 1, 'MD', 'Radiology', 220, 1);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (26, 'Emily Wilson', 'Emily', 'Wilson', 1, 'Surgeon', 678901234, 1, 'MD', 'Surgery', 230, 0);

INSERT INTO dbo.users(emplid, fullname, fname, lname, status, position, ssn, eligible, title, dept, depnum, deleted)
VALUES (27, 'John Lee', 'John', 'Lee', 1, 'Urologist', 789012345, 1, 'MD', 'Urology', 240, 1);


SELECT * FROM [].[] WHERE eligible = '1' AND deleted = '0'

