import psycopg2
import xlsxwriter
import psycopg2.extras

class FilmScript:
    def __init__(self):

	    self.workbook = xlsxwriter.Workbook('Filmsandactor.xlsx')

	    self.worksheet = self.workbook.add_worksheet("FilmandActor")
	    self.worksheet.write_row('A1', ["Film title", "actor First name", "A last name", "F language"])
	    self.row = 1
	    self.col = 0

    def TestConnection(self):
        self.connection = psycopg2.connect(user = "nasir",
                                        password = "your password",
                                        host = "localhost",
                                        port = "5432",
                                        database = "database name"  
                                        )


        self.cursor = self.connection.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        self.cursor.execute("""SELECT  actor_id, film.film_id , language.language_id ,film_actor.last_update return_date from film_actor  INNER JOIN film on film_actor.film_id = film.film_id 
            inner join language on film.language_id = language.language_id """) 


        film_actor = self.cursor.fetchall()
        for films_actor in film_actor:
            self.cursor.execute(
                """SELECT title from film where film_id =%s""", (films_actor['film_id'],))
            film_title = self.cursor.fetchone()
            self.cursor.execute(
                """SELECT first_name, last_name from actor where actor_id =%s""", (films_actor['actor_id'],))
            actor_all = self.cursor.fetchone()
            self.cursor.execute(
                """ SELECT name from language where language_id =%s """, (films_actor['language_id'],)
            )
            language_id = self.cursor.fetchone()
            if language_id['name'] == 'English':

                print film_title['title'], actor_all['first_name'],actor_all['last_name'],language_id['name']
                self.worksheet.write_row(self.row, self.col, [film_title['title'], actor_all['first_name'], actor_all['last_name'], language_id['name'] ])
                self.row += 1
            else:
                print("noting to Fetch")    
obj = FilmScript()
obj.TestConnection()
obj.workbook.close()
