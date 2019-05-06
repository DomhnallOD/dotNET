using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        static string conString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\dotNET_Project\dotNET_Assignment_7.accdb";
        OleDbConnection connection = new OleDbConnection(conString);
        OleDbCommand cmd;
        OleDbDataAdapter adapter;
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable dt4 = new DataTable();
        DataTable dt5 = new DataTable();
        DataTable dt6 = new DataTable();
        public Form1()
        {
            InitializeComponent();
            //DataGridView1 Properties
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].Name = "Client ID";
            dataGridView1.Columns[1].Name = "Client Name";
            dataGridView1.Columns[2].Name = "Date of Birth";
            dataGridView1.Columns[3].Name = "Client Address";

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //Selection Mode
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;


            //DataGridView1 Properties
            dataGridView2.ColumnCount = 5;
            dataGridView2.Columns[0].Name = "Contract ID";
            dataGridView2.Columns[1].Name = "Contract Holder";
            dataGridView2.Columns[2].Name = "Contract Begins";
            dataGridView2.Columns[3].Name = "Contract Ends";
            dataGridView2.Columns[4].Name = "Value";

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //Selection Mode
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.MultiSelect = false;

        }

        //Insert into DB
        private void add(string name, DateTime dob, string address)
        {
            //SQL statement
            string sql = "INSERT INTO Clients(Client_Name, Client_DOB, Client_Address) VALUES(@name, @dob, @address)";
            cmd = new OleDbCommand(sql, connection);

            //Add Parameters
            cmd.Parameters.AddWithValue("@name", name);
            cmd.Parameters.AddWithValue("@dob", dob);
            cmd.Parameters.AddWithValue("@address", address);

            try
            {
                connection.Open();
                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("Record inserted successfully.");
                }
                clearTextBoxes();
                connection.Close();
                retrieve();
            } catch (Exception ex)
            {
                MessageBox.Show("Error has occurred: Please input values for new record creation.");
                connection.Close();
            }
        }


        //Fill DataGridView1
        private void populate(string id, string name, DateTime dob, string address)
        {
            dataGridView1.Rows.Add(id, name, dob, address);
        }

        //Fill DataGridView2
        private void populateContracts(string id, string holder, DateTime begins, DateTime ends, string value)
        {
            dataGridView2.Rows.Add(id, holder, begins, ends, value);
        }

        //Retrieve from DB
        private void retrieve()
        {
            //RETRIEVE DATA FOR CLIENTS TABLE
            dataGridView1.Rows.Clear();
            //SQL statement
            String sql = "SELECT * FROM Clients";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt);

                //Loop through contents of dt
                foreach (DataRow row in dt.Rows)
                {
                    populate(row[0].ToString(), row[1].ToString(), Convert.ToDateTime(row[2].ToString()), row[3].ToString());
                }
                clearTextBoxes();
                connection.Close();
                //Clear DT 
                dt.Rows.Clear();
            } catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }


            //RETRIEVE DATA FOR CONTRACTS TABLE
            dataGridView2.Rows.Clear();
            //SQL Statement
            sql = "SELECT * FROM Contracts";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt2);

                //Loop through contents of dt
                foreach (DataRow row in dt2.Rows)
                {
                    populateContracts(row[0].ToString(), row[1].ToString(), Convert.ToDateTime(row[2].ToString()), Convert.ToDateTime(row[3].ToString()), row[4].ToString());
                }
                clearTextBoxes();
                connection.Close();
                //Clear DT 
                dt2.Rows.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }

        }

        //Update DB
        private void update(int id, string name, DateTime dob, string address)
        {
            //SQL statement
            string sql = "Update Clients SET Client_Name='" + name + "', Client_DOB=" + dob.ToShortDateString() + ",Client_Address='" + address + "' WHERE Client_ID=" + id;
            cmd = new OleDbCommand(sql, connection);

            //Open connection, update, retrieve, dataGridView
            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.UpdateCommand = connection.CreateCommand();
                adapter.UpdateCommand.CommandText = sql;

                if (adapter.UpdateCommand.ExecuteNonQuery() > 0)
                {
                    clearTextBoxes();
                    MessageBox.Show("Successfully updated row.");
                }
                clearTextBoxes();
                connection.Close();

                //Refresh
                retrieve();
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

        //Delete from DB
        private void delete(int id)
        {
            //SQL statement
            string sql = "DELETE FROM Clients WHERE Client_ID = " + id + "";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.DeleteCommand = connection.CreateCommand();
                adapter.DeleteCommand.CommandText = sql;

                //Prompt for confirmation
                if(MessageBox.Show("Are you sure?", "Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        clearTextBoxes();
                        MessageBox.Show("Row has been successfuly deleted.");
                    }
                }

                connection.Close();
                retrieve();
            } catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex);
                connection.Close();
            }
        }

        private void contractsPerClient(int id)
        {
            //SQL statement
            string sql = "SELECT Client_Name, COUNT(Contracts.Contract_ID) FROM Contracts LEFT JOIN Clients  ON Clients.Client_ID = Contracts.Contract_Holder WHERE Clients.Client_ID=" + id + " GROUP BY Clients.Client_Name";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt3);

                //Loop through contents of dt
                string result = "";
                int countRows = 0;
                if (dt4.Rows.Count > 0)
                {
                    foreach (DataRow row in dt3.Rows)
                    {
                        countRows += 1;
                        result += "Client name: " + row[0].ToString() + " \t Number of Contracts: " + row[1].ToString() + "\n";
                    }
                    connection.Close();

                    MessageBox.Show(result.ToString());
                    //Clear DT 
                    dt3.Rows.Clear();
                } else
                {
                    MessageBox.Show("Client is yet to sign a contract.");
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }
        }

        //Average contract value per client
        private void valuePerClient(int id)
        {
            //SQL statement
            string sql = "SELECT Clients.Client_Name, AVG(Contracts.Value) FROM Contracts LEFT JOIN Clients  ON Clients.Client_ID = Contracts.Contract_Holder WHERE Clients.Client_ID=" + id + " GROUP BY Clients.Client_Name";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt3);

                //Loop through contents of dt
                string result = "";
                int countRows = 0;
                if (dt3.Rows.Count > 0)
                {
                    foreach (DataRow row in dt3.Rows)
                    {
                        countRows += 1;
                        result += "Client name: " + row[0].ToString() + " \t Average contract Value: €" + row[1].ToString() + "\n";
                    }
                    connection.Close();

                    MessageBox.Show(result.ToString());
                    //Clear DT 
                    dt3.Rows.Clear();
                } else
                {
                    MessageBox.Show("Client is yet to sign a contract.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }
        }

        //Average contract duration per client
        private void averageContractDuration(int id)
        {
            //SQL statement
            string sql = "SELECT Clients.Client_Name, ROUND(AVG(DATEDIFF('d', Contracts.Contract_Begins, Contracts.Contract_Ends))) AS DateDiff FROM Contracts LEFT JOIN Clients ON Clients.Client_ID = Contracts.Contract_Holder WHERE Clients.Client_ID=" + id + " GROUP BY Clients.Client_Name";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt4);

                //Loop through contents of dt
                string result = "";
                int countRows = 0;
                if (dt4.Rows.Count > 0)
                {
                    foreach (DataRow row in dt4.Rows)
                    {
                        countRows += 1;
                        result += "Client name: " + row[0].ToString() + " \t Average contract duration: " + row[1].ToString() + " days\n";
                    }
                    connection.Close();

                    MessageBox.Show(result.ToString());
                    //Clear DT 
                    dt4.Rows.Clear();
                    connection.Close();
                } else
                {
                    connection.Close();

                    MessageBox.Show("Client is yet to sign a contract.");
                    //Clear DT 
                    dt4.Rows.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }
        }

        //Time remaining on a specific contract
        private void timeRemainingOnContract(int id)
        {
            //SQL statement
            string sql = "SELECT Contracts.Contract_ID, Clients.Client_Name, ROUND(DATEDIFF('d', DATE(), Contracts.Contract_Ends)) AS DateDiff FROM Contracts LEFT JOIN Clients ON Clients.Client_ID = Contracts.Contract_Holder WHERE Contracts.Contract_ID=" + id;
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt5);

                //Loop through contents of dt
                string result = "";
                int countRows = 0;
                foreach (DataRow row in dt5.Rows)
                {
                    countRows += 1;
                    result += "Contract ID: " + row[0].ToString() + "\tClient name: " + row[1].ToString() + " \tTime remaining on contract: " + row[2].ToString() + " days\n";
                }
                connection.Close();

                MessageBox.Show(result.ToString());
                //Clear DT 
                dt5.Rows.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }
        }

        //Number of contracts currently open
        private void contractsOpen()
        {
            //SQL statement
            string sql = "SELECT COUNT(Contracts.Contract_ID) AS 'Open Contracts' FROM Contracts LEFT JOIN Clients ON Clients.Client_ID = Contracts.Contract_Holder WHERE DATEDIFF('d', DATE(), Contracts.Contract_Ends) > 0";
            cmd = new OleDbCommand(sql, connection);

            try
            {
                connection.Open();
                adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(dt6);

                //Loop through contents of dt
                string result = "";
                int countRows = 0;
                foreach (DataRow row in dt6.Rows)
                {
                    countRows += 1;
                    result += "Number of open contracts: " + row[0].ToString();
                }
                connection.Close();

                MessageBox.Show(result.ToString());
                //Clear DT 
                dt6.Rows.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                connection.Close();
            }
        }

        //Clear text boxes
        private void clearTextBoxes()
        {
            textBox1.Text = "";
            dateTimePicker1.Value = DateTime.Now;
            textBox2.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            add(textBox1.Text, dateTimePicker1.Value, textBox2.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            retrieve();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                int id = Convert.ToInt32(selected);

                update(id, textBox1.Text, dateTimePicker1.Value, textBox2.Text);
            }catch(Exception ex)
            {
                MessageBox.Show("Error has occurred: Please select which of the client records you wish to update, and input the new values");
            }
        }

        private void button4_Click(object sender, EventArgs e) {
            string selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            delete(id);
            
        }
        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            clearTextBoxes();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            contractsPerClient(id);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            valuePerClient(id);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            averageContractDuration(id);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string selected = dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            timeRemainingOnContract(id);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            contractsOpen();
        }
    }
}
