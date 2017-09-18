<%@ Page Title="" Trace="true" Language="C#" MasterPageFile="~/MasterPage1.master" AutoEventWireup="true" CodeFile="tests.aspx.cs" Inherits="tests" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Word套版測試" />
    <hr />
    <asp:TextBox ID="TextBox1" runat="server" Height="100px" Rows="5" TextMode="MultiLine" Width="438px"></asp:TextBox>
    <br />
    <asp:Button ID="Button2" runat="server" Text="測試" OnClick="Button2_Click" />
</asp:Content>

