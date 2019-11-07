# yc.excel
导出excel支持列拆分

## example

``` JAVA
@Data
public class Person{
  private String firstName;
  private String lastName;
  private Integer age;
}

public void export(HttpServletResponse response){
   ExcelDocument doc = new ExcelDocument();
   
   ExcelSheet sheet = doc.addSheet("sheet1");
   
   //data init
   List<Person> persons = lists.newArrayList();
   ....
   
   //cell init
   List<ExcelCol> cols = lists.newArrayList();
   cols.add(
      new TreeCol("name")
          .child(new DataCol<Person>("fistName").value(t->t.getFirstName())
          .child(new DataCol<Person>("lastName").value(t->t.getLastName())
   );
   cols.add(new DataCol<Person>("age").value(t->t.getAge()));
   
   sheet.write(persons,cols);
   
   doc.writeToResponse(response,"demo");
}
```

> output:
<table>
<tr>
<td colSpan="2">name</td>
<td>age</td>
</tr>
<tr>
<td>peter</td>
<td>lin</td>
<td>23</td>
</tr>
</table>
