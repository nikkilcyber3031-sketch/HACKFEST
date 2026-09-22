from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_from_directory,
    send_file
)

import os
import uuid

from werkzeug.utils import secure_filename
from database import get_db_connection

from openpyxl import Workbook
from io import BytesIO


# =========================================================
# FLASK APP
# =========================================================

app = Flask(__name__)


# =========================================================
# UPLOAD SETTINGS
# =========================================================

UPLOAD_FOLDER = "uploads/payment_screenshots"

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {
    "png",
    "jpg",
    "jpeg"
}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# =========================================================
# FILE VALIDATION
# =========================================================

def allowed_file(filename):

    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


# =========================================================
# PAGES
# =========================================================

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/register")
def register():
    return render_template("register.html")


@app.route("/payment")
def payment():
    return render_template("payment.html")


@app.route("/admin-login")
def admin_login():
    return render_template("admin-login.html")


@app.route("/admin-dashboard")
def admin_dashboard():
    return render_template("admin-dashboard.html")


# =========================================================
# TEST DATABASE
# =========================================================

@app.route("/test-db")
def test_db():

    connection = None
    cursor = None

    try:

        connection = get_db_connection()

        cursor = connection.cursor()

        cursor.execute("SELECT DATABASE()")

        database_name = cursor.fetchone()[0]

        return f"Database Connected Successfully: {database_name}"

    except Exception as e:

        return f"Database Connection Failed: {e}"

    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# REGISTER TEAM API
# =========================================================

@app.route("/api/register", methods=["POST"])
def register_team():

    connection = None
    cursor = None

    try:

        data = request.get_json()

        if not data:

            return jsonify({
                "success": False,
                "message": "No registration data received."
            }), 400


        # -------------------------------------------------
        # TEAM DETAILS
        # -------------------------------------------------

        team_name = data.get(
            "teamName",
            ""
        ).strip()

        college_name = data.get(
            "collegeName",
            ""
        ).strip()

        department = data.get(
            "department",
            ""
        ).strip()

        team_size = data.get("teamSize")

        members = data.get(
            "members",
            []
        )


        # -------------------------------------------------
        # BASIC VALIDATION
        # -------------------------------------------------

        if not team_name or not college_name or not department:

            return jsonify({
                "success": False,
                "message": "Please fill all team details."
            }), 400


        try:

            team_size = int(team_size)

        except:

            return jsonify({
                "success": False,
                "message": "Invalid team size."
            }), 400


        if team_size not in [4, 5, 6]:

            return jsonify({
                "success": False,
                "message": "Team size must be 4, 5 or 6."
            }), 400


        if len(members) != team_size:

            return jsonify({
                "success": False,
                "message": "Member count does not match team size."
            }), 400


        # -------------------------------------------------
        # DATABASE CONNECTION
        # -------------------------------------------------

        connection = get_db_connection()

        cursor = connection.cursor()


        # -------------------------------------------------
        # INSERT TEAM
        # -------------------------------------------------

        team_query = """
            INSERT INTO teams
            (
                team_name,
                college_name,
                department,
                team_size
            )
            VALUES (%s, %s, %s, %s)
        """


        cursor.execute(
            team_query,
            (
                team_name,
                college_name,
                department,
                team_size
            )
        )


        team_id = cursor.lastrowid


        # -------------------------------------------------
        # INSERT MEMBERS
        # -------------------------------------------------

        member_query = """
            INSERT INTO members
            (
                team_id,
                name,
                email,
                phone,
                is_leader
            )
            VALUES (%s, %s, %s, %s, %s)
        """


        for index, member in enumerate(members):

            name = member.get(
                "name",
                ""
            ).strip()

            email = member.get(
                "email",
                ""
            ).strip()

            phone = member.get(
                "phone",
                ""
            ).strip()


            # -------------------------------------------------
            # MEMBER VALIDATION
            # -------------------------------------------------

            if not name or not email or not phone:

                connection.rollback()

                return jsonify({
                    "success": False,
                    "message":
                    f"Member {index + 1} details are incomplete."
                }), 400


            # First member = Team Leader

            is_leader = index == 0


            cursor.execute(
                member_query,
                (
                    team_id,
                    name,
                    email,
                    phone,
                    is_leader
                )
            )


        # -------------------------------------------------
        # SAVE REGISTRATION
        # -------------------------------------------------

        connection.commit()


        return jsonify({

            "success": True,

            "message":
            "Registration saved successfully.",

            "team_id":
            team_id

        })


    except Exception as e:

        if connection:
            connection.rollback()

        return jsonify({

            "success": False,

            "message":
            str(e)

        }), 500


    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# PAYMENT API
# =========================================================

@app.route("/api/payment", methods=["POST"])
def submit_payment():

    connection = None
    cursor = None

    try:

        team_id = request.form.get(
            "team_id"
        )

        transaction_id = request.form.get(
            "transaction_id",
            ""
        ).strip()

        screenshot = request.files.get(
            "screenshot"
        )


        # -------------------------------------------------
        # VALIDATION
        # -------------------------------------------------

        if not team_id:

            return jsonify({
                "success": False,
                "message": "Team ID is missing."
            }), 400


        if not transaction_id:

            return jsonify({
                "success": False,
                "message": "Transaction ID is required."
            }), 400


        if not screenshot:

            return jsonify({
                "success": False,
                "message":
                "Payment screenshot is required."
            }), 400


        if not allowed_file(
            screenshot.filename
        ):

            return jsonify({
                "success": False,
                "message":
                "Only JPG, JPEG or PNG files are allowed."
            }), 400


        # -------------------------------------------------
        # DATABASE
        # -------------------------------------------------

        connection = get_db_connection()

        cursor = connection.cursor()


        # -------------------------------------------------
        # CHECK TEAM
        # -------------------------------------------------

        cursor.execute(
            """
            SELECT team_size
            FROM teams
            WHERE id = %s
            """,
            (team_id,)
        )


        team = cursor.fetchone()


        if not team:

            return jsonify({
                "success": False,
                "message": "Team not found."
            }), 404


        team_size = int(
            team[0]
        )


        # -------------------------------------------------
        # CALCULATE PAYMENT
        # -------------------------------------------------

        amount = team_size * 200


        # -------------------------------------------------
        # SAVE SCREENSHOT
        # -------------------------------------------------

        original_name = secure_filename(
            screenshot.filename
        )


        extension = original_name.rsplit(
            ".",
            1
        )[1].lower()


        unique_name = (
            str(uuid.uuid4())
            + "."
            + extension
        )


        file_path = os.path.join(
            app.config["UPLOAD_FOLDER"],
            unique_name
        )


        screenshot.save(
            file_path
        )


        # -------------------------------------------------
        # SAVE PAYMENT
        # -------------------------------------------------

        payment_query = """
            INSERT INTO payments
            (
                team_id,
                amount,
                transaction_id,
                screenshot,
                payment_status
            )
            VALUES (%s, %s, %s, %s, %s)
        """


        cursor.execute(
            payment_query,
            (
                team_id,
                amount,
                transaction_id,
                unique_name,
                "Pending Verification"
            )
        )


        connection.commit()


        return jsonify({

            "success": True,

            "message":
            "Payment submitted successfully.",

            "team_id":
            team_id,

            "amount":
            amount,

            "transaction_id":
            transaction_id,

            "payment_status":
            "Pending Verification"

        })


    except Exception as e:

        if connection:
            connection.rollback()

        return jsonify({

            "success": False,

            "message":
            str(e)

        }), 500


    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# PAYMENT SCREENSHOT
# =========================================================

@app.route(
    "/payment-screenshot/<filename>"
)
def payment_screenshot(filename):

    return send_from_directory(

        app.config["UPLOAD_FOLDER"],

        filename

    )


# =========================================================
# ADMIN TEAMS API
# =========================================================

@app.route("/api/admin/teams")
def admin_teams():

    connection = None
    cursor = None

    try:

        connection = get_db_connection()

        cursor = connection.cursor(
            dictionary=True
        )


        # -------------------------------------------------
        # GET ALL TEAMS
        # -------------------------------------------------

        cursor.execute("""
            SELECT
                id,
                team_name,
                college_name,
                department,
                team_size,
                created_at
            FROM teams
            ORDER BY id DESC
        """)


        teams = cursor.fetchall()


        # -------------------------------------------------
        # GET MEMBERS AND PAYMENT
        # -------------------------------------------------

        for team in teams:


            # MEMBERS

            cursor.execute("""
                SELECT
                    id,
                    name,
                    email,
                    phone,
                    is_leader
                FROM members
                WHERE team_id = %s
                ORDER BY id
            """, (
                team["id"],
            ))


            team["members"] = (
                cursor.fetchall()
            )


            # PAYMENT

            cursor.execute("""
                SELECT
                    amount,
                    transaction_id,
                    screenshot,
                    payment_status,
                    submitted_at
                FROM payments
                WHERE team_id = %s
                ORDER BY id DESC
                LIMIT 1
            """, (
                team["id"],
            ))


            team["payment"] = (
                cursor.fetchone()
            )


        return jsonify({

            "success": True,

            "teams": teams

        })


    except Exception as e:

        return jsonify({

            "success": False,

            "message":
            str(e)

        }), 500


    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# UPDATE PAYMENT STATUS
# =========================================================

@app.route(
    "/api/admin/payment-status",
    methods=["POST"]
)
def update_payment_status():

    connection = None
    cursor = None

    try:

        data = request.get_json()


        team_id = data.get(
            "team_id"
        )

        status = data.get(
            "status"
        )


        if (
            not team_id
            or status not in [
                "Verified",
                "Rejected"
            ]
        ):

            return jsonify({

                "success": False,

                "message":
                "Invalid team ID or payment status"

            }), 400


        connection = get_db_connection()

        cursor = connection.cursor()


        cursor.execute("""
            UPDATE payments
            SET payment_status = %s
            WHERE team_id = %s
            ORDER BY id DESC
            LIMIT 1
        """, (
            status,
            team_id
        ))


        connection.commit()


        if cursor.rowcount == 0:

            return jsonify({

                "success": False,

                "message":
                "Payment not found"

            }), 404


        return jsonify({

            "success": True,

            "message":
            f"Payment {status.lower()} successfully"

        })


    except Exception as e:

        if connection:
            connection.rollback()

        return jsonify({

            "success": False,

            "message":
            str(e)

        }), 500


    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# DOWNLOAD EXCEL
# =========================================================

@app.route(
    "/admin/download-excel"
)
def download_excel():

    connection = None
    cursor = None

    try:

        connection = get_db_connection()

        cursor = connection.cursor(
            dictionary=True
        )


        # -------------------------------------------------
        # GET ALL TEAMS
        # -------------------------------------------------

        cursor.execute("""
            SELECT
                id,
                team_name,
                college_name,
                department
            FROM teams
            ORDER BY id
        """)


        teams = cursor.fetchall()


        # -------------------------------------------------
        # CREATE EXCEL
        # -------------------------------------------------

        workbook = Workbook()

        worksheet = workbook.active

        worksheet.title = (
            "HACKFEST Registrations"
        )


        # -------------------------------------------------
        # HEADINGS
        # -------------------------------------------------

        worksheet.append([

            "Team Name",

            "College Name",

            "Department",

            "Member Name",

            "Email",

            "Phone",

            "Role"

        ])


        # -------------------------------------------------
        # ADD TEAM DATA
        # -------------------------------------------------

        for team in teams:


            # Get members

            cursor.execute("""
                SELECT
                    name,
                    email,
                    phone,
                    is_leader
                FROM members
                WHERE team_id = %s
                ORDER BY id
            """, (
                team["id"],
            ))


            members = cursor.fetchall()


            # Add members

            for member in members:


                if str(
                    member["is_leader"]
                ) in [
                    "1",
                    "True",
                    "true"
                ]:

                    role = "Team Leader"

                else:

                    role = "Member"


                worksheet.append([

                    team["team_name"],

                    team["college_name"],

                    team["department"],

                    member["name"],

                    member["email"],

                    member["phone"],

                    role

                ])


            # Blank row between teams

            worksheet.append([])


        # -------------------------------------------------
        # COLUMN WIDTH
        # -------------------------------------------------

        worksheet.column_dimensions[
            "A"
        ].width = 25

        worksheet.column_dimensions[
            "B"
        ].width = 30

        worksheet.column_dimensions[
            "C"
        ].width = 20

        worksheet.column_dimensions[
            "D"
        ].width = 25

        worksheet.column_dimensions[
            "E"
        ].width = 35

        worksheet.column_dimensions[
            "F"
        ].width = 18

        worksheet.column_dimensions[
            "G"
        ].width = 18


        # -------------------------------------------------
        # SAVE EXCEL IN MEMORY
        # -------------------------------------------------

        output = BytesIO()

        workbook.save(
            output
        )

        output.seek(0)


        # -------------------------------------------------
        # DOWNLOAD
        # -------------------------------------------------

        return send_file(

            output,

            as_attachment=True,

            download_name=
            "HACKFEST_Registrations.xlsx",

            mimetype=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )


    except Exception as e:

        return f"""
        <h2>Excel Download Error</h2>
        <p>{str(e)}</p>
        """, 500


    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# CLEAR ALL REGISTRATION TEST DATA
# =========================================================

@app.route(
    "/api/admin/clear-data",
    methods=["POST"]
)
def clear_registration_data():

    connection = None
    cursor = None

    try:

        connection = get_db_connection()

        cursor = connection.cursor()


        # -------------------------------------------------
        # DELETE CHILD DATA FIRST
        # -------------------------------------------------

        cursor.execute(
            "DELETE FROM payments"
        )


        cursor.execute(
            "DELETE FROM members"
        )


        cursor.execute(
            "DELETE FROM teams"
        )


        # -------------------------------------------------
        # SAVE CHANGES
        # -------------------------------------------------

        connection.commit()


        return jsonify({

            "success": True,

            "message":
            "All registration test data cleared successfully."

        })


    except Exception as e:

        if connection:
            connection.rollback()

        return jsonify({

            "success": False,

            "message":
            str(e)

        }), 500


    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# =========================================================
# ADMIN TEST API
# =========================================================

@app.route("/api/admin/test")
def admin_test():

    return jsonify({

        "success": True,

        "message":
        "Admin API route is working"

    })


# =========================================================
# RUN SERVER
# =========================================================

if __name__ == "__main__":

    app.run(

        debug=True,

        host="127.0.0.1",

        port=5000

    )