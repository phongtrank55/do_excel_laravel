<!DOCTYPE html>
<html lang="en">
<head>
  <title>Do Excel</title>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
</head>
<body>

<div class="container">
  <h2>Xuất excel</h2>
  <div class="pull-right">
        <form action="{{ route('export') }}" method="POST">
            @csrf
            <button class="btn btn-primary" type="submit">Xuất</button>
        </form>
  </div>
  <table class="table">
    <thead>
      <tr>
        <th>#</th>
        <th>Tên</th>
        <th>Lastname</th>
        <th>Mô tả</th>
        <th>Thời gian</th>
      </tr>
    </thead>
    <tbody>
        @php
            $i = 1;   
        @endphp
        @foreach ($products as $product)
            <tr>
                <td>{{ $i++ }}</td>
                <td>{{ $product->name }}</td>
                <td>{{ $product->description }}</td>
                <td>{{ $product->quantity }}</td>
                <td>
                    <p><span class="text-bold">Ngày tạo: </span>{{ $product->created_at }} </p>
                    <p><span class="text-bold">Cập nhật: </span>{{ $product->updated_at }} </p>
                </td>
            </tr>
        @endforeach
    </tbody>
  </table>
</div>

</body>
</html>
